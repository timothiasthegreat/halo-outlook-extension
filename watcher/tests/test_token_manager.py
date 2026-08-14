"""Tests for TokenManager — per-user Halo OAuth token lifecycle."""

import tempfile
from unittest.mock import AsyncMock, patch

import pytest

from watcher.crypto import encrypt_token, generate_key
from watcher.state import StateStore
from watcher.token_manager import TokenManager


@pytest.fixture
def fernet_key():
    return generate_key()


@pytest.fixture
async def state(fernet_key):
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


def make_token_manager(fernet_key, state, **overrides):
    """Create a TokenManager with test config."""
    config = {
        "instance_url": "https://test.halopsa.com",
        "client_id": "test-client-id",
        "client_secret": "",  # PKCE — no client secret needed
    }
    config.update(overrides)
    return TokenManager(config, state, fernet_key)


def mock_token_response(status=200, access_token="new-access-token", refresh_token="rt-xyz-new", expires_in=3600):
    """Create a mock httpx response for the token endpoint.

    httpx.Response.json() is a synchronous method, not a coroutine.
    """
    from unittest.mock import MagicMock
    resp = AsyncMock()
    resp.status_code = status
    resp.json = MagicMock(return_value={
        "access_token": access_token,
        "refresh_token": refresh_token,
        "expires_in": expires_in,
        "token_type": "Bearer",
    })
    return resp


@pytest.mark.asyncio
class TestTokenManagerRegistration:
    async def test_register_user_stores_encrypted_token(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        plain_token = "refresh-token-abc-123"

        await tm.register_user("user@example.com", plain_token)

        row = await state.get_mailbox("user@example.com")
        assert row is not None
        assert row["refresh_token_enc"] != plain_token
        assert row["token_status"] == "active"

    async def test_register_user_decrypts_correctly(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        plain_token = "my-refresh-token-value"

        await tm.register_user("user@example.com", plain_token)

        row = await state.get_mailbox("user@example.com")
        from watcher.crypto import decrypt_token
        decrypted = decrypt_token(row["refresh_token_enc"], fernet_key)
        assert decrypted == plain_token


@pytest.mark.asyncio
class TestTokenManagerGetToken:
    async def test_get_token_for_registered_user_calls_token_endpoint(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        await tm.register_user("user@example.com", "rt-abc")

        mock_resp = mock_token_response(access_token="at-1")
        with patch("httpx.AsyncClient.post", new_callable=AsyncMock) as mock_post:
            mock_post.return_value = mock_resp

            access_token = await tm.get_token_for("user@example.com")

            assert access_token == "at-1"
            assert mock_post.call_count == 1
            call_data = mock_post.call_args[1]["data"]
            assert call_data["grant_type"] == "refresh_token"
            assert call_data["refresh_token"] == "rt-abc"

    async def test_get_token_for_unregistered_user_raises(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)

        with pytest.raises(ValueError, match="not registered"):
            await tm.get_token_for("nobody@example.com")

    async def test_get_token_uses_cache_within_expiry_window(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        await tm.register_user("user@example.com", "rt-abc")

        mock_resp = mock_token_response(access_token="first-token")
        with patch("httpx.AsyncClient.post", new_callable=AsyncMock) as mock_post:
            mock_post.return_value = mock_resp

            token1 = await tm.get_token_for("user@example.com")
            token2 = await tm.get_token_for("user@example.com")

            assert token1 == "first-token"
            assert token2 == "first-token"
            assert mock_post.call_count == 1  # cached second call

    async def test_expired_token_marks_mailbox_expired(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        await tm.register_user("user@example.com", "rt-abc")

        mock_resp = mock_token_response(status=400)
        mock_resp.json.return_value = {"error": "invalid_grant"}
        with patch("httpx.AsyncClient.post", new_callable=AsyncMock) as mock_post:
            mock_post.return_value = mock_resp

            with pytest.raises(Exception):
                await tm.get_token_for("user@example.com")

            row = await state.get_mailbox("user@example.com")
            assert row["token_status"] == "expired"


@pytest.mark.asyncio
class TestTokenManagerRefresh:
    async def test_token_refresh_updates_stored_refresh_token(self, fernet_key, state):
        tm = make_token_manager(fernet_key, state)
        await tm.register_user("user@example.com", "old-rt")

        mock_resp = mock_token_response(access_token="at-2", refresh_token="new-rt-from-halo")
        with patch("httpx.AsyncClient.post", new_callable=AsyncMock) as mock_post:
            mock_post.return_value = mock_resp

            await tm.get_token_for("user@example.com")

            row = await state.get_mailbox("user@example.com")
            from watcher.crypto import decrypt_token
            decrypted = decrypt_token(row["refresh_token_enc"], fernet_key)
            assert decrypted == "new-rt-from-halo"