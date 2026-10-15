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
    mock.list_actions = AsyncMock(return_value=[])
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

    async def test_sync_skips_message_already_journaled_on_ticket(self, state, token_manager):
        """The add-in journals the first email itself; the watcher must not repost it."""
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-dup", 4242, watched_by="user@example.com")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Original email",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "internetMessageId": "<first@test.com>",
                },
                {
                    "subject": "Reply",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "internetMessageId": "<second@test.com>",
                },
            ]
        )

        mock_halo = make_mock_halo_client()
        mock_halo.list_actions = AsyncMock(
            return_value=[
                {"id": 1, "email_message_id": "<first@test.com>"},
                {"id": 2, "email_message_id": None},
            ]
        )
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        posted = [c.kwargs["email_message_id"] for c in mock_halo.create_action.call_args_list]
        assert posted == ["<second@test.com>"]
        assert mock_halo.list_actions.await_count == 1
        # Recorded locally so later cycles don't ask Halo again
        assert await state.is_message_synced("conv-dup", "<first@test.com>")

    async def test_sync_does_not_query_halo_when_all_synced(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()

        await state.track_conversation("conv-quiet", 4343, watched_by="user@example.com")
        await state.mark_synced("conv-quiet", "<old@test.com>")

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Old",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "internetMessageId": "<old@test.com>",
                }
            ]
        )

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            await engine.sync_once()

        mock_halo.list_actions.assert_not_awaited()
        mock_halo.create_action.assert_not_awaited()

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

@pytest.mark.asyncio
class TestAttachmentMailbox:
    async def test_attachments_read_from_watching_mailbox(self, state, token_manager):
        """Graph message ids only resolve in their own mailbox, so attachments
        must come from watched_by, not graph.user_email."""
        config = make_config()
        graph = AsyncMock()
        await state.track_conversation("conv-att", 9999, watched_by="user@example.com")
        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "id": "graph-msg-1",
                    "subject": "With file",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "internetMessageId": "<att@test.com>",
                    "hasAttachments": True,
                }
            ]
        )
        graph.get_message_attachments = AsyncMock(
            return_value=[
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "att-1",
                    "name": "report.pdf",
                    "contentType": "application/pdf",
                }
            ]
        )
        graph.get_attachment_content = AsyncMock(return_value=b"%PDF")

        mock_halo = make_mock_halo_client()
        mock_halo.attach_to_ticket = AsyncMock()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            await SyncEngine(config, token_manager, graph, state).sync_once()

        graph.get_message_attachments.assert_awaited_once_with(
            "graph-msg-1", user_email="user@example.com"
        )
        graph.get_attachment_content.assert_awaited_once_with(
            "graph-msg-1", "att-1", user_email="user@example.com"
        )
        mock_halo.attach_to_ticket.assert_awaited_once()


@pytest.mark.asyncio
class TestStaleConversations:
    def _late_reply(self):
        return [
            {
                "subject": "Late reply",
                "from": {"emailAddress": {"address": "customer@example.com"}},
                "internetMessageId": "<late@test.com>",
            }
        ]

    async def test_stale_conversation_still_journals_late_reply(self, state, token_manager):
        config = make_config()
        graph = AsyncMock()
        await state.track_conversation("conv-quiet", 9999, watched_by="user@example.com")
        await state.mark_stale("conv-quiet")
        graph.get_messages_by_conversation = AsyncMock(return_value=self._late_reply())

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            stats = await SyncEngine(config, token_manager, graph, state).sync_once()

        assert stats["messages_synced"] == 1
        mock_halo.create_action.assert_awaited_once()
        # The reply reactivates the conversation
        watched = await state.get_watched_conversations()
        assert [c["conversation_id"] for c in watched] == ["conv-quiet"]
        assert await state.get_stale_conversations() == []

    async def test_stale_conversations_polled_less_often(self, state, token_manager):
        config = make_config(watcher={"stale_poll_interval_seconds": 3600})
        graph = AsyncMock()
        await state.track_conversation("conv-quiet", 9999, watched_by="user@example.com")
        await state.mark_stale("conv-quiet")
        graph.get_messages_by_conversation = AsyncMock(return_value=[])

        mock_halo = make_mock_halo_client()
        with patch("watcher.sync.HaloClient", return_value=mock_halo):
            engine = SyncEngine(config, token_manager, graph, state)
            first = await engine.sync_once()
            second = await engine.sync_once()
            engine._last_stale_poll -= 3600  # interval elapsed
            third = await engine.sync_once()

        assert first["conversations_checked"] == 1
        assert second["conversations_checked"] == 0
        assert third["conversations_checked"] == 1

    async def test_quiet_conversation_marked_stale_not_unlinked(self, state, token_manager):
        config = make_config(watcher={"stale_conversation_days": 1})
        await state.track_conversation("conv-old", 9999, watched_by="user@example.com")
        await state._conn.execute(
            "UPDATE conversations SET last_sync_at = '2000-01-01T00:00:00+00:00'"
        )
        await state._conn.commit()

        graph = AsyncMock()
        graph.get_messages_by_conversation = AsyncMock(return_value=[])
        with patch("watcher.sync.HaloClient", return_value=make_mock_halo_client()):
            await SyncEngine(config, token_manager, graph, state).sync_once()

        assert await state.get_ticket_id("conv-old") == 9999
        assert [c["conversation_id"] for c in await state.get_stale_conversations()] == ["conv-old"]
