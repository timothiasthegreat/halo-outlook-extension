"""Tests for the watcher's internal API (/register-mailbox, /mailbox-status)."""

from unittest.mock import AsyncMock, MagicMock

from fastapi.testclient import TestClient

from watcher.token_manager import (
    RegistrationMismatchError,
    TokenRefreshError,
    TokenRevokedError,
)
from watcher.watcher import _build_health_app


def make_client(token_manager=None, state=None):
    shared_state = {"_token_manager": token_manager, "_state_store": state}
    return TestClient(_build_health_app(shared_state))


def test_register_mailbox_uses_verified_registration():
    tm = MagicMock()
    tm.register_verified = AsyncMock()
    client = make_client(token_manager=tm)

    resp = client.post(
        "/register-mailbox", json={"email": "User@Example.com", "refresh_token": "rt"}
    )

    assert resp.status_code == 200
    assert resp.json()["watching"] == ["user@example.com"]
    tm.register_verified.assert_awaited_once_with("user@example.com", "rt")


def test_register_mailbox_ignores_additional_emails():
    tm = MagicMock()
    tm.register_verified = AsyncMock()
    client = make_client(token_manager=tm)

    resp = client.post(
        "/register-mailbox",
        json={"email": "a@example.com", "refresh_token": "rt", "additional_emails": ["b@x.com"]},
    )

    assert resp.status_code == 200
    assert tm.register_verified.await_count == 1


def test_register_mailbox_mismatch_returns_403():
    tm = MagicMock()
    tm.register_verified = AsyncMock(side_effect=RegistrationMismatchError("not yours"))
    client = make_client(token_manager=tm)

    resp = client.post("/register-mailbox", json={"email": "v@example.com", "refresh_token": "rt"})

    assert resp.status_code == 403


def test_register_mailbox_rejected_token_returns_400():
    tm = MagicMock()
    tm.register_verified = AsyncMock(side_effect=TokenRevokedError("bad"))
    client = make_client(token_manager=tm)

    resp = client.post("/register-mailbox", json={"email": "v@example.com", "refresh_token": "rt"})

    assert resp.status_code == 400


def test_register_mailbox_halo_down_returns_502():
    tm = MagicMock()
    tm.register_verified = AsyncMock(side_effect=TokenRefreshError("503"))
    client = make_client(token_manager=tm)

    resp = client.post("/register-mailbox", json={"email": "v@example.com", "refresh_token": "rt"})

    assert resp.status_code == 502


def test_mailbox_status_reports_token_state():
    state = MagicMock()
    state.get_mailbox = AsyncMock(return_value={"email": "a@example.com", "token_status": "expired"})
    client = make_client(state=state)

    resp = client.get("/mailbox-status", params={"email": "A@example.com"})

    assert resp.status_code == 200
    assert resp.json()["registered"] is True
    assert resp.json()["token_status"] == "expired"
    state.get_mailbox.assert_awaited_once_with("a@example.com")


def test_mailbox_status_unregistered():
    state = MagicMock()
    state.get_mailbox = AsyncMock(return_value=None)
    client = make_client(state=state)

    resp = client.get("/mailbox-status", params={"email": "nobody@example.com"})

    assert resp.json() == {"email": "nobody@example.com", "registered": False, "token_status": None}


# ── /conversation and /untrack-conversation ─────────────────────────────


def _conv_row(watched_by="agent@example.com", is_stale=0):
    return {
        "conversation_id": "conv-1",
        "ticket_id": 42,
        "created_at": "2026-10-01T00:00:00+00:00",
        "last_sync_at": "2026-10-02T00:00:00+00:00",
        "is_stale": is_stale,
        "watched_by": watched_by,
    }


def test_get_conversation_returns_mapping_for_watching_mailbox():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=_conv_row(is_stale=1))
    client = make_client(state=state)

    resp = client.get(
        "/conversation", params={"conversation_id": "conv-1", "watched_by": "Agent@Example.com"}
    )

    assert resp.status_code == 200
    assert resp.json() == {
        "conversationId": "conv-1",
        "ticketId": 42,
        "lastSyncAt": "2026-10-02T00:00:00+00:00",
        "isStale": True,
    }


def test_get_conversation_hides_other_mailboxes_mapping():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=_conv_row(watched_by="other@example.com"))
    client = make_client(state=state)

    resp = client.get(
        "/conversation", params={"conversation_id": "conv-1", "watched_by": "agent@example.com"}
    )

    assert resp.status_code == 404


def test_get_conversation_untracked_returns_404():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=None)
    client = make_client(state=state)

    resp = client.get(
        "/conversation", params={"conversation_id": "nope", "watched_by": "agent@example.com"}
    )

    assert resp.status_code == 404


def test_untrack_conversation_unlinks():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=_conv_row())
    state.unlink_conversation = AsyncMock()
    client = make_client(state=state)

    resp = client.post(
        "/untrack-conversation",
        json={"conversationId": "conv-1", "watchedBy": "Agent@example.com"},
    )

    assert resp.status_code == 200
    state.unlink_conversation.assert_awaited_once_with("conv-1")


def test_untrack_conversation_watched_by_someone_else_returns_403():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=_conv_row(watched_by="other@example.com"))
    state.unlink_conversation = AsyncMock()
    client = make_client(state=state)

    resp = client.post(
        "/untrack-conversation",
        json={"conversationId": "conv-1", "watchedBy": "agent@example.com"},
    )

    assert resp.status_code == 403
    state.unlink_conversation.assert_not_awaited()


def test_untrack_conversation_untracked_returns_404():
    state = MagicMock()
    state.get_conversation = AsyncMock(return_value=None)
    state.unlink_conversation = AsyncMock()
    client = make_client(state=state)

    resp = client.post(
        "/untrack-conversation",
        json={"conversationId": "conv-1", "watchedBy": "agent@example.com"},
    )

    assert resp.status_code == 404
    state.unlink_conversation.assert_not_awaited()
