"""Tests for SQLite state store."""

import tempfile

import pytest

from watcher.state import StateStore


@pytest.fixture
async def store():
    """Create a StateStore with a temp in-memory database."""
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.mark.asyncio
class TestStateStore:
    async def test_track_and_get_conversation(self, store):
        await store.track_conversation("conv-001", 1234)
        ticket_id = await store.get_ticket_id("conv-001")
        assert ticket_id == 1234

    async def test_untracked_conversation_returns_none(self, store):
        ticket_id = await store.get_ticket_id("nonexistent")
        assert ticket_id is None

    async def test_stale_conversation_returns_none(self, store):
        await store.track_conversation("conv-001", 1234)
        await store.mark_stale("conv-001")
        ticket_id = await store.get_ticket_id("conv-001")
        assert ticket_id is None

    async def test_get_watched_conversations(self, store):
        await store.track_conversation("conv-001", 100)
        await store.track_conversation("conv-002", 200)
        await store.mark_stale("conv-002")  # Should not appear

        watched = await store.get_watched_conversations()
        assert len(watched) == 1
        assert watched[0]["conversation_id"] == "conv-001"
        assert watched[0]["ticket_id"] == 100

    async def test_track_updates_existing(self, store):
        await store.track_conversation("conv-001", 100)
        await store.track_conversation("conv-001", 200)  # Update ticket ID
        ticket_id = await store.get_ticket_id("conv-001")
        assert ticket_id == 200

    async def test_unlink_removes_everything(self, store):
        await store.track_conversation("conv-001", 1234)
        await store.mark_synced("conv-001", "msg-1@test.com")
        await store.unlink_conversation("conv-001")
        assert await store.get_ticket_id("conv-001") is None
        assert not await store.is_message_synced("conv-001", "msg-1@test.com")

    async def test_touch_sync_updates_timestamp(self, store):
        await store.track_conversation("conv-001", 1234)
        watched_before = await store.get_watched_conversations()
        before = watched_before[0]["last_sync_at"]

        import asyncio

        await asyncio.sleep(0.01)  # Ensure timestamp difference
        await store.touch_sync("conv-001")

        watched_after = await store.get_watched_conversations()
        after = watched_after[0]["last_sync_at"]
        assert after != before


@pytest.mark.asyncio
class TestMessageDedup:
    async def test_message_synced_false_initially(self, store):
        assert not await store.is_message_synced("conv-001", "msg-1@test.com")

    async def test_mark_and_check_synced(self, store):
        await store.mark_synced("conv-001", "msg-1@test.com")
        assert await store.is_message_synced("conv-001", "msg-1@test.com")

    async def test_duplicate_mark_returns_false(self, store):
        first = await store.mark_synced("conv-001", "msg-1@test.com")
        assert first is True

        second = await store.mark_synced("conv-001", "msg-1@test.com")
        assert second is False

    async def test_same_message_different_conversation(self, store):
        """Same messageId in different conversations should be separate."""
        await store.mark_synced("conv-001", "msg-1@test.com")
        assert not await store.is_message_synced("conv-002", "msg-1@test.com")

    async def test_prune_stale_synced(self, store):
        await store.mark_synced("conv-001", "msg-old@test.com")
        await store.mark_synced("conv-001", "msg-new@test.com")

        # prune with 0 days — everything older than "now" should be pruned
        # but there's a sub-second timing issue; use a generous cutoff
        count = await store.prune_stale_synced(-1)  # negative days = future cutoff
        # Both should be pruned since they were all added before "future"
        assert count == 2
        assert not await store.is_message_synced("conv-001", "msg-old@test.com")
        assert not await store.is_message_synced("conv-001", "msg-new@test.com")