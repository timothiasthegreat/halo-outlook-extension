"""Per-user Halo OAuth2 token manager.

Manages the lifecycle of per-user refresh tokens: loading from state.db,
decrypting, requesting access tokens via PKCE refresh_token grant, caching,
and persisting new refresh tokens when Halo rotates them.

Unlike the old HaloClient._acquire_token() which used client_credentials,
this manager uses per-user refresh tokens. No client secret is needed —
PKCE-originated refresh tokens authenticate with just client_id + refresh_token.

The watcher holds its own refresh token family: the add-in runs a separate
authorization for it, so a rotation on one side never invalidates the token
the other side holds.

Usage:
    tm = TokenManager(halo_config_dict, state_store, fernet_key)
    await tm.register_verified("user@example.com", "plaintext-refresh-token")
    access_token = await tm.get_token_for("user@example.com")
"""

from __future__ import annotations

import time

import httpx
import structlog

from watcher.crypto import decrypt_token, encrypt_token
from watcher.state import StateStore

logger = structlog.get_logger()


class TokenRevokedError(Exception):
    """Halo rejected the refresh token (HTTP 400/401) — the user must re-register."""


class TokenRefreshError(Exception):
    """Transient token endpoint failure (5xx, 429, ...) — retry next cycle."""


class RegistrationMismatchError(Exception):
    """The refresh token belongs to a different Halo agent than the claimed email."""


class TokenManager:
    """Manages per-user Halo OAuth access tokens.

    Loads encrypted refresh tokens from state.db, decrypts them, exchanges
    them for access tokens at Halo's /auth/token endpoint, caches results
    until expiry, and persists rotated refresh tokens back to state.db.
    """

    def __init__(
        self,
        halo_config: dict,
        state: StateStore,
        fernet_key: str,
    ):
        self._instance_url = halo_config["instance_url"].rstrip("/")
        self._client_id = halo_config["client_id"]
        self._state = state
        self._fernet_key = fernet_key

        # Token cache: email → (access_token, expiry_ts)
        self._cache: dict[str, tuple[str, float]] = {}

    @property
    def token_url(self) -> str:
        return f"{self._instance_url}/auth/token"

    # ── registration ───────────────────────────────────────────

    async def register_user(self, email: str, refresh_token: str) -> None:
        """Encrypt and store a refresh token for a new or existing user.

        The token is encrypted with Fernet before storage. If the user is
        already registered, this updates their token and resets the expired
        status — effectively re-registering. Does not check who owns the
        token; use register_verified() for tokens received over the network.
        """
        enc = encrypt_token(refresh_token, self._fernet_key)
        await self._state.register_mailbox(email, enc)
        self._cache.pop(email, None)
        logger.info("user_registered", email=email)

    async def register_verified(self, email: str, refresh_token: str) -> None:
        """Register a refresh token only if it belongs to the Halo agent ``email``.

        Exchanges the token, asks Halo who it belongs to (/api/agent/me) and
        stores the (possibly rotated) refresh token.

        Raises TokenRevokedError if Halo rejects the token, TokenRefreshError
        on transient failures, RegistrationMismatchError if the token belongs
        to a different agent.
        """
        data = await self._refresh(refresh_token)
        access_token = data["access_token"]

        async with httpx.AsyncClient(timeout=15.0) as client:
            response = await client.get(
                f"{self._instance_url}/api/agent/me",
                headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json"},
            )
        if response.status_code != 200:
            raise TokenRefreshError(f"agent/me returned HTTP {response.status_code}")
        agent_email = str(response.json().get("email") or "").strip().lower()
        if agent_email != email.strip().lower():
            logger.warning("registration_email_mismatch", email=email, agent_email=agent_email)
            raise RegistrationMismatchError(
                f"Halo token belongs to {agent_email or 'an agent with no email'}, not {email}"
            )

        stored_rt = data.get("refresh_token") or refresh_token
        await self._state.register_mailbox(email, encrypt_token(stored_rt, self._fernet_key))
        self._cache[email] = (access_token, time.monotonic() + data.get("expires_in", 3600))
        logger.info("user_registered", email=email, verified=True)

    # ── token acquisition ──────────────────────────────────────

    async def get_token_for(self, email: str) -> str:
        """Get a valid access token for the given user.

        Returns a cached token if still within the expiry window (60s buffer).
        Otherwise decrypts the stored refresh token, calls Halo's token
        endpoint with the refresh_token grant, caches the result, and
        persists any rotated refresh token.

        Raises ValueError if the user is not registered.
        Raises TokenRevokedError if Halo rejects the token (and marks it expired).
        Raises TokenRefreshError / httpx.HTTPError on transient failures; the
        token stays active and is retried next cycle.
        """
        # Check cache
        cached = self._cache.get(email)
        if cached:
            token, expires = cached
            if time.monotonic() < expires - 60:
                return token

        # Load from state.db
        row = await self._state.get_mailbox(email)
        if row is None or row["token_status"] != "active":
            raise ValueError(f"User {email} is not registered or token is expired")

        # Decrypt and exchange
        try:
            refresh_token = decrypt_token(str(row["refresh_token_enc"]), self._fernet_key)
        except Exception:
            await self._state.mark_token_expired(email)
            raise ValueError(f"Failed to decrypt refresh token for {email}")

        try:
            data = await self._refresh(refresh_token)
        except TokenRevokedError:
            await self._state.mark_token_expired(email)
            raise

        access_token = data["access_token"]
        self._cache[email] = (access_token, time.monotonic() + data.get("expires_in", 3600))

        # Persist rotated refresh token if present
        new_rt = data.get("refresh_token")
        if new_rt and new_rt != refresh_token:
            await self._state.update_token(email, encrypt_token(new_rt, self._fernet_key))
            logger.debug("token_rotated", email=email)

        return access_token

    async def _refresh(self, refresh_token: str) -> dict:
        """Run the refresh_token grant and return the token response body."""
        async with httpx.AsyncClient(timeout=15.0) as client:
            response = await client.post(
                self.token_url,
                data={
                    "grant_type": "refresh_token",
                    "client_id": self._client_id,
                    "refresh_token": refresh_token,
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )

        if response.status_code == 200:
            return response.json()

        logger.error(
            "token_refresh_failed",
            status=response.status_code,
            body=response.text[:200],
        )
        # 400 (invalid_grant) and 401 mean the token itself was rejected.
        # Anything else (429, 5xx) is a transient Halo problem.
        if response.status_code in (400, 401):
            raise TokenRevokedError(f"Token refresh rejected: HTTP {response.status_code}")
        raise TokenRefreshError(f"Token refresh failed: HTTP {response.status_code}")
