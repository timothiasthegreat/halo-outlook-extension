"""End-to-end test: full sync pipeline with multiple conversations."""

import asyncio
import tempfile
from unittest.mock import AsyncMock, call

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
class TestEndToEnd:
    async def test_full_pipeline_multiple_conversations(self, state):
        """Simulate: 3 conversations, 2-4 messages each, mixed directions."""
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        # Track 3 conversations
        await state.track_conversation("conv-e2e-1", 1001)
        await state.track_conversation("conv-e2e-2", 1002)
        await state.track_conversation("conv-e2e-3", 1003)

        # Pre-sync one message in conv-1 (it should be skipped via dedup)
        await state.mark_synced("conv-e2e-1", "<conv1-already-synced@test.com>")

        # Mock Graph responses per conversation
        messages_map = {
            "conv-e2e-1": [
                {
                    "subject": "Already synced",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-10T10:00:00Z",
                    "internetMessageId": "<conv1-already-synced@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>Old</p>"},
                },
                {
                    "subject": "New inbound",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "internetMessageId": "<conv1-new@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>New customer message</p>"},
                },
            ],
            "conv-e2e-2": [
                {
                    "subject": "Agent reply",
                    "from": {"emailAddress": {"address": "tim@firesideit.ca"}},
                    "sentDateTime": "2026-08-11T14:00:00Z",
                    "internetMessageId": "<conv2-outbound@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>Hi from Tim</p>"},
                },
                {
                    "subject": "Customer reply",
                    "from": {"emailAddress": {"address": "cust@example.com"}},
                    "sentDateTime": "2026-08-12T09:00:00Z",
                    "internetMessageId": "<conv2-inbound@test.com>",
                    "hasAttachments": True,
                    "id": "graph-msg-004",
                    "body": {"content": "<p>Response with attachment</p>"},
                },
            ],
            "conv-e2e-3": [
                {
                    "subject": "Single message",
                    "from": {"emailAddress": {"address": "vendor@external.com"}},
                    "sentDateTime": "2026-08-12T15:00:00Z",
                    "internetMessageId": "<conv3-inbound@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>Vendor update</p>"},
                },
                {
                    "subject": "Tim reply",
                    "from": {"emailAddress": {"address": "tim@firesideit.ca"}},
                    "sentDateTime": "2026-08-12T15:05:00Z",
                    "internetMessageId": "<conv3-outbound@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>Thanks</p>"},
                },
                {
                    "subject": "Another vendor note",
                    "from": {"emailAddress": {"address": "vendor@external.com"}},
                    "sentDateTime": "2026-08-12T15:10:00Z",
                    "internetMessageId": "<conv3-inbound2@test.com>",
                    "hasAttachments": False,
                    "body": {"content": "<p>More info</p>"},
                },
                {
                    "subject": "No internetMessageId",
                    "from": {"emailAddress": {"address": "no-id@test.com"}},
                    "sentDateTime": "2026-08-12T15:15:00Z",
                    "hasAttachments": False,
                    "body": {"content": "<p>No ID</p>"},
                },
            ],
        }

        async def get_messages(conv_id: str):
            return messages_map.get(conv_id, [])

        graph.get_messages_by_conversation = AsyncMock(side_effect=get_messages)
        graph.get_message_attachments = AsyncMock(
            return_value=[
                {
                    "id": "att-005",
                    "name": "document.docx",
                    "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "@odata.type": "#microsoft.graph.fileAttachment",
                }
            ]
        )
        graph.get_attachment_content = AsyncMock(return_value=b"fake-docx")

        halo.create_action = AsyncMock(return_value={"id": 1})
        halo.attach_to_ticket = AsyncMock(return_value={"id": 200})

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        # Expected:
        # conv-1: 1 already synced (skip) + 1 new = 1
        # conv-2: 2 new messages = 2
        # conv-3: 4 messages, 1 with no internetMessageId (skip) = 3
        # Total: 1 + 2 + 3 = 6
        assert stats["conversations_checked"] == 3
        assert stats["messages_synced"] == 6
        assert stats["errors"] == 0

        # Verify outcome IDs are correct
        calls = halo.create_action.call_args_list
        outcome_ids = [c.kwargs["outcome_id"] for c in calls]

        # Count by direction
        sent_count = sum(1 for oid in outcome_ids if oid == 16)  # email_sent
        received_count = sum(1 for oid in outcome_ids if oid == 0)  # email_received
        assert sent_count == 2, f"Expected 2 outbound actions, got {sent_count}"
        assert received_count == 4, f"Expected 4 inbound actions, got {received_count}"

        # Verify dedup
        assert await state.is_message_synced(
            "conv-e2e-1", "<conv1-already-synced@test.com>"
        )
        assert await state.is_message_synced(
            "conv-e2e-1", "<conv1-new@test.com>"
        )

        # Attachment was uploaded for conv-2's customer reply
        assert halo.attach_to_ticket.call_count == 1

    async def test_message_without_internet_id_skipped(self, state):
        """Edge case: messages without internetMessageId should not crash."""
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        await state.track_conversation("conv-no-msgid", 9000)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=[
                {
                    "subject": "No ID",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "hasAttachments": False,
                    "body": {"content": "<p>Test</p>"},
                    # No internetMessageId
                }
            ]
        )

        halo.create_action = AsyncMock()

        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()

        assert stats["messages_synced"] == 0
        halo.create_action.assert_not_called()