"""Tests for watched_mailboxes state management.

The watched_mailboxes table persists per-user refresh tokens (encrypted at rest)
and drives the watcher's multi-mailbox sync loop.
"""

import tempfile

import pytest

from watcher.state import StateStore


@pytest.fixture
async def store():
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.mark.asyncio
class TestWatchedMailboxes:
    async def test_register_and_get_mailbox(self, store):
        await store.register_mailbox(
            "user@example.com", "encrypted-refresh-token-123"
        )
        row = await store.get_mailbox("user@example.com")
        assert row is not None
        assert row["email"] == "user@example.com"
        assert row["refresh_token_enc"] == "encrypted-refresh-token-123"
        assert row["token_status"] == "active"

    async def test_get_mailbox_nonexistent_returns_none(self, store):
        row = await store.get_mailbox("nobody@example.com")
        assert row is None

    async def test_registered_mailbox_appears_in_watched_list(self, store):
        await store.register_mailbox("a@example.com", "tok-a")
        await store.register_mailbox("b@example.com", "tok-b")

        watched = await store.get_watched_mailboxes()
        assert len(watched) == 2
        emails = {r["email"] for r in watched}
        assert emails == {"a@example.com", "b@example.com"}

    async def test_stale_mailboxes_excluded_from_watched_list(self, store):
        await store.register_mailbox("a@example.com", "tok-a")
        await store.register_mailbox("b@example.com", "tok-b")
        await store.register_mailbox("c@example.com", "tok-c")

        # Mark one as expired
        await store.mark_token_expired("b@example.com")

        watched = await store.get_watched_mailboxes()
        assert len(watched) == 2
        emails = {r["email"] for r in watched}
        assert emails == {"a@example.com", "c@example.com"}

    async def test_update_token_replaces_stored_token(self, store):
        await store.register_mailbox("user@example.com", "old-token")
        await store.update_token("user@example.com", "new-token")

        row = await store.get_mailbox("user@example.com")
        assert row["refresh_token_enc"] == "new-token"
        assert row["token_status"] == "active"
        assert row["last_token_refresh_at"] is not None

    async def test_update_token_resets_expired_status(self, store):
        await store.register_mailbox("user@example.com", "old-token")
        await store.mark_token_expired("user@example.com")

        # Now update the token — should go back to active
        await store.update_token("user@example.com", "fresh-token")
        row = await store.get_mailbox("user@example.com")
        assert row["token_status"] == "active"

    async def test_mark_token_expired(self, store):
        await store.register_mailbox("user@example.com", "tok")
        await store.mark_token_expired("user@example.com")

        row = await store.get_mailbox("user@example.com")
        assert row["token_status"] == "expired"

    async def test_mark_expired_nonexistent_no_error(self, store):
        """Marking a nonexistent mailbox should not raise."""
        await store.mark_token_expired("ghost@example.com")
        # No exception = pass

    async def test_unregister_removes_row(self, store):
        await store.register_mailbox("user@example.com", "tok")
        await store.unregister_mailbox("user@example.com")

        row = await store.get_mailbox("user@example.com")
        assert row is None

    async def test_unregister_removes_from_watched_list(self, store):
        await store.register_mailbox("a@example.com", "tok-a")
        await store.register_mailbox("b@example.com", "tok-b")
        await store.unregister_mailbox("a@example.com")

        watched = await store.get_watched_mailboxes()
        assert len(watched) == 1
        assert watched[0]["email"] == "b@example.com"

    async def test_reregister_updates_existing(self, store):
        """Re-registering the same email should update, not duplicate."""
        await store.register_mailbox("user@example.com", "first-token")
        await store.mark_token_expired("user@example.com")
        await store.register_mailbox("user@example.com", "second-token")

        row = await store.get_mailbox("user@example.com")
        assert row["refresh_token_enc"] == "second-token"
        assert row["token_status"] == "active"

        # Should still be only one row
        watched = await store.get_watched_mailboxes()
        assert len(watched) == 1

    async def test_register_mailbox_with_status_active(self, store):
        """Default status on register should be 'active'."""
        await store.register_mailbox("new@example.com", "tok")
        row = await store.get_mailbox("new@example.com")
        assert row["token_status"] == "active"