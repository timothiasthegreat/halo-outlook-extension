"""Tests for attachment sync and staleness detection."""

import asyncio
import tempfile
from datetime import datetime, timedelta, timezone
from unittest.mock import AsyncMock, patch

import pytest

from watcher.config import Config
from watcher.sync import SyncEngine
from watcher.state import StateStore


def make_config(stale_days: int = 14) -> Config:
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
            "watcher": {
                "stale_conversation_days": stale_days,
            },
        }
    )


@pytest.fixture
async def state():
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.mark.asyncio
class TestAttachmentSync:
    async def test_attachments_fetched_and_uploaded(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-att", 5000)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "With attachment",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "internetMessageId": "<msg-att@test.com>",
                    "hasAttachments": True,
                    "id": "graph-msg-id-001",
                    "body": {"content": "<p>See attached</p>"},
                }
            ]
        )
        graph.get_message_attachments = AsyncMock(
            return_value=[
                {
                    "id": "att-001",
                    "name": "report.pdf",
                    "contentType": "application/pdf",
                    "size": 1024,
                    "@odata.type": "#microsoft.graph.fileAttachment",
                }
            ]
        )
        graph.get_attachment_content = AsyncMock(return_value=b"fake-pdf-content")

        halo.create_action = AsyncMock(return_value={"id": 100})
        halo.attach_to_ticket = AsyncMock(return_value={"id": 200})

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        assert halo.create_action.call_count == 1
        assert halo.attach_to_ticket.call_count == 1

        attach_call = halo.attach_to_ticket.call_args
        assert attach_call.kwargs["ticket_id"] == 5000
        assert attach_call.kwargs["filename"] == "report.pdf"
        assert attach_call.kwargs["content_type"] == "application/pdf"
        assert attach_call.kwargs["content"] == b"fake-pdf-content"

    async def test_attachments_skip_inline_images(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-inline", 5001)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "With inline image",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "internetMessageId": "<msg-inline@test.com>",
                    "hasAttachments": True,
                    "id": "graph-msg-id-002",
                    "body": {"content": "<p>See image</p>"},
                }
            ]
        )
        graph.get_message_attachments = AsyncMock(
            return_value=[
                {
                    "id": "item-att",
                    "name": "inline.png",
                    "@odata.type": "#microsoft.graph.itemAttachment",
                },
                {
                    "id": "ref-att",
                    "name": "reference",
                    "@odata.type": "#microsoft.graph.referenceAttachment",
                },
            ]
        )

        halo.create_action = AsyncMock(return_value={"id": 101})
        halo.attach_to_ticket = AsyncMock()

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        assert halo.attach_to_ticket.call_count == 0

    async def test_attachments_oversized_skipped(self, state):
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-big", 5002)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "Big file",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "internetMessageId": "<msg-big@test.com>",
                    "hasAttachments": True,
                    "id": "graph-msg-id-003",
                    "body": {"content": "<p>Big file</p>"},
                }
            ]
        )
        graph.get_message_attachments = AsyncMock(
            return_value=[
                {
                    "id": "att-big",
                    "name": "large.iso",
                    "contentType": "application/octet-stream",
                    "@odata.type": "#microsoft.graph.fileAttachment",
                }
            ]
        )
        graph.get_attachment_content = AsyncMock(
            return_value=b"x" * (11 * 1024 * 1024)
        )

        halo.create_action = AsyncMock(return_value={"id": 102})
        halo.attach_to_ticket = AsyncMock()

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        assert stats["messages_synced"] == 1
        assert halo.attach_to_ticket.call_count == 0


@pytest.mark.asyncio
class TestStalenessDetection:
    async def test_stale_conversations_marked(self, state):
        config = make_config(stale_days=1)
        halo = AsyncMock()
        graph = AsyncMock()

        # Insert a conversation with last_sync_at 30 days ago
        old_date = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
        await state.track_conversation("conv-old", 100)
        assert state._conn is not None
        await state._conn.execute(
            "UPDATE conversations SET last_sync_at = ? WHERE conversation_id = ?",
            (old_date, "conv-old"),
        )
        await state._conn.commit()

        # Fresh conversation
        await state.track_conversation("conv-fresh", 200)

        graph.get_messages_by_conversation = AsyncMock(return_value=[])

        engine = SyncEngine(config, halo, graph, state)
        await engine.sync_once()

        assert await state.get_ticket_id("conv-old") is None  # stale
        assert await state.get_ticket_id("conv-fresh") == 200  # still active

    async def test_stale_logs_emitted(self, state):
        config = make_config(stale_days=1)
        halo = AsyncMock()
        graph = AsyncMock()

        old_date = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
        await state.track_conversation("conv-stale-log", 300)
        assert state._conn is not None
        await state._conn.execute(
            "UPDATE conversations SET last_sync_at = ? WHERE conversation_id = ?",
            (old_date, "conv-stale-log"),
        )
        await state._conn.commit()

        graph.get_messages_by_conversation = AsyncMock(return_value=[])

        with patch("watcher.sync.logger") as mock_log:
            engine = SyncEngine(config, halo, graph, state)
            await engine.sync_once()

            stale_logged = any(
                "stale_conversations_marked" in str(call)
                for call in mock_log.info.call_args_list
            )
            assert stale_logged, "Expected stale_conversations_marked log"