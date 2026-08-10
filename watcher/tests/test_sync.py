"""Tests for the SyncEngine orchestration."""

import asyncio
import tempfile
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from watcher.config import Config
from watcher.sync import SyncEngine
from watcher.state import StateStore


def make_config() -> Config:
    return Config.model_validate(
        {
            "halo": {
                "instance_url": "https://test.halopsa.com",
                "client_id": "test-client",
                "client_secret": "test-secret",
            },
            "graph": {
                "tenant_id": "test-tenant",
                "client_id": "test-graph-client",
                "client_secret": "test-graph-secret",
                "user_email": "tim@firesideit.ca",
            },
        }
    )


@pytest.fixture
async def state():
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.mark.asyncio
class TestSyncEngine:
    async def test_sync_skips_no_watched_conversations(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()
        engine = SyncEngine(config, halo, graph, state)

        stats = await engine.sync_once()

        assert stats["conversations_checked"] == 0
        assert stats["messages_synced"] == 0

    async def test_sync_dedup_skips_already_synced(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        # Track a conversation
        await state.track_conversation("conv-123", 9999)
        # Pre-mark a message as synced
        await state.mark_synced("conv-123", "<msg-already-synced@test.com>")

        # Graph returns two messages — one new, one already synced
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
        halo.create_action = AsyncMock(
            return_value={"id": 1, "ticket_id": 9999, "outcome_id": 0}
        )

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        # Only the new message should be synced
        assert stats["messages_synced"] == 1
        # Halo action called once for the new message
        assert halo.create_action.call_count == 1

        # Verify the already-synced is still marked synced
        assert await state.is_message_synced("conv-123", "<msg-already-synced@test.com>")
        # Verify the new message is now marked synced
        assert await state.is_message_synced("conv-123", "<msg-new@test.com>")

    async def test_sync_determines_direction_correctly(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-456", 1000)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Outbound from Tim",
                    "from": {"emailAddress": {"address": "tim@firesideit.ca"}},
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
        halo.create_action = AsyncMock(return_value={"id": 1})

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        assert stats["messages_synced"] == 2
        assert halo.create_action.call_count == 2

        # Check outcome_ids: outbound → email_sent (16), inbound → email_received (0)
        calls = halo.create_action.call_args_list
        outcome_ids = {call.kwargs["outcome_id"] for call in calls}
        assert outcome_ids == {0, 16}

    async def test_sync_marks_conversation_touched(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-789", 2000)

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
        halo.create_action = AsyncMock(return_value={"id": 5})

        engine = SyncEngine(config, halo, graph, state)
        await engine.sync_once()

        # Verify conversation is still active (not stale)
        watched = await state.get_watched_conversations()
        assert len(watched) == 1
        assert watched[0]["conversation_id"] == "conv-789"

    async def test_sync_error_on_one_conversation_continues(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-good", 100)
        await state.track_conversation("conv-bad", 200)

        # Good conv returns messages
        def get_messages(conversation_id):
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
        halo.create_action = AsyncMock(return_value={"id": 1})

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        # conv-good synced 1, conv-bad errored
        assert stats["messages_synced"] == 1
        assert stats["errors"] == 1
        assert stats["conversations_checked"] == 2

    async def test_messages_without_internet_id_skipped(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-no-id", 3000)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "No ID",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T12:00:00Z",
                    # No internetMessageId
                    "body": {"content": "<p>No ID</p>"},
                }
            ]
        )
        halo.create_action = AsyncMock()

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        # Should not attempt to create action
        assert stats["messages_synced"] == 0
        halo.create_action.assert_not_called()