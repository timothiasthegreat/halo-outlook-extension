"""Tests for the SyncEngine with multi-mailbox support."""

import asyncio
import tempfile
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from watcher.config import Config
from watcher.crypto import generate_key
from watcher.sync import SyncEngine
from watcher.state import StateStore
from watcher.token_manager import TokenManager


def make_config(**overrides):
    cfg = {
        "halo": {
            "instance_url": "https://test.halopsa.com",
            "client_id": "test-client",
            "client_secret": "test-secret",
        },
        "graph": {
            "tenant_id": "test-tenant",
            "client_id": "test-graph-client",
            "client_secret": "test-graph-secret",
            "user_email": "default@example.com",
        },
    }
    for key, val in overrides.items():
        cfg[key] = val
    return Config.model_validate(cfg)


@pytest.fixture
async def state():
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.fixture
def fernet_key():
    return generate_key()


@pytest.fixture
async def token_manager(fernet_key):
    """Create a TokenManager with a cached test user."""
    async with StateStore(":memory:") as state:
        tm = TokenManager(
            {"instance_url": "https://test.halopsa.com", "client_id": "test-client"},
            state,
            fernet_key,
        )
        await tm.register_user("user@example.com", "dummy-rt")
        tm._cache["user@example.com"] = ("mock-access-token", float("inf"))
        yield tm


def make_mock_halo_client():
    """Create a mock HaloClient for sync tests."""
    mock = AsyncMock()
    mock.__aenter__ = AsyncMock(return_value=mock)
    mock.__aexit__ = AsyncMock(return_value=None)
    mock.create_action = AsyncMock(return_value={"id": 1, "ticket_id": 9999, "outcome_id": 0})
    return mock


@pytest.mark.asyncio
class TestSyncEngine:
    async def test_sync_skips_no_watched_conversations(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()
        engine = SyncEngine(config, token_manager, graph, state)

        stats = await engine.sync_once()

        assert stats["conversations_checked"] == 0
        assert stats["messages_synced"] == 0

    async def test_sync_dedup_skips_already_synced(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-123", 9999, watched_by="user@example.com")
        await state.mark_synced("conv-123", "<msg-already-synced@test.com>")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Already synced",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:00:00Z",
                    "internetMessageId": "<msg-already-synced@test.com>",
                    "conversationId": "conv-123",
                },
                {
                    "subject": "New message",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:05:00Z",
                    "internetMessageId": "<msg-new@test.com>",
                    "conversationId": "conv-123",
                    "body": {"content": "<p>New</p>"},
                },
            ]
        )

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        assert await state.is_message_synced("conv-123", "<msg-already-synced@test.com>")
        assert await state.is_message_synced("conv-123", "<msg-new@test.com>")

    async def test_sync_determines_direction_correctly(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-456", 1000, watched_by="user@example.com")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Outbound from Tim",
                    "from": {"emailAddress": {"address": "user@example.com"}},
                    "sentDateTime": "2026-08-10T13:00:00Z",
                    "internetMessageId": "<outbound@test.com>",
                    "body": {"content": "<p>Hi</p>"},
                },
                {
                    "subject": "Inbound from customer",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:55:00Z",
                    "internetMessageId": "<inbound@test.com>",
                    "body": {"content": "<p>Question</p>"},
                },
            ]
        )

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            stats = await engine.sync_once()

        assert stats["messages_synced"] == 2

    async def test_sync_marks_conversation_touched(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-789", 2000, watched_by="user@example.com")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "New",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:00:00Z",
                    "internetMessageId": "<msg-single@test.com>",
                    "body": {"content": "<p>Test</p>"},
                }
            ]
        )

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            await engine.sync_once()

        watched = await state.get_watched_conversations()
        assert len(watched) == 1
        assert watched[0]["conversation_id"] == "conv-789"

    async def test_sync_error_on_one_conversation_continues(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-good", 100, watched_by="user@example.com")
        await state.track_conversation("conv-bad", 200, watched_by="user@example.com")

        def get_messages(conversation_id, **kwargs):
            if conversation_id == "conv-good":
                return [
                    {
                        "subject": "Good msg",
                        "from": {"emailAddress": {"address": "customer@example.com"}},
                        "sentDateTime": "2026-08-10T12:00:00Z",
                        "internetMessageId": "<good@test.com>",
                        "body": {"content": "<p>Good</p>"},
                    }
                ]
            else:
                raise Exception("Graph API down for this conversation!")

        graph.get_messages_by_conversation = AsyncMock(side_effect=get_messages)

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        assert stats["errors"] == 1
        assert stats["conversations_checked"] == 2

    async def test_conversation_without_watcher_skipped(self, state, token_manager):
        """Conversations without watched_by should be skipped gracefully."""
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-no-owner", 9999)  # no watched_by

        engine = SyncEngine(config, token_manager, graph, state)
        stats = await engine.sync_once()

        assert stats["conversations_checked"] == 1
        assert stats["messages_synced"] == 0
        graph.get_messages_by_conversation.assert_not_called()

    async def test_messages_without_internet_id_skipped(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-no-id", 3000, watched_by="user@example.com")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "No ID",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:00:00Z",
                    "body": {"content": "<p>No ID</p>"},
                }
            ]
        )

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            stats = await engine.sync_once()

        assert stats["messages_synced"] == 0