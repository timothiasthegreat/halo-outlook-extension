"""Per-user Halo OAuth2 token manager.

Manages the lifecycle of per-user refresh tokens: loading from state.db,
decrypting, requesting access tokens via PKCE refresh_token grant, caching,
and persisting new refresh tokens when Halo rotates them.

Unlike the old HaloClient._acquire_token() which used client_credentials,
this manager uses per-user refresh tokens. No client secret is needed —
PKCE-originated refresh tokens authenticate with just client_id + refresh_token.

Usage:
    tm = TokenManager(halo_config_dict, state_store, fernet_key)
    await tm.register_user("user@example.com", "plaintext-refresh-token")
    access_token = await tm.get_token_for("user@example.com")
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field

import httpx
import structlog

from watcher.state import StateStore
from watcher.crypto import decrypt_token, encrypt_token

logger = structlog.get_logger()


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
        status — effectively re-registering.
        """
        enc = encrypt_token(refresh_token, self._fernet_key)
        await self._state.register_mailbox(email, enc)
        logger.info("user_registered", email=email)

    # ── token acquisition ──────────────────────────────────────

    async def get_token_for(self, email: str) -> str:
        """Get a valid access token for the given user.

        Returns a cached token if still within the expiry window (60s buffer).
        Otherwise decrypts the stored refresh token, calls Halo's token
        endpoint with the refresh_token grant, caches the result, and
        persists any rotated refresh token.

        Raises ValueError if the user is not registered.
        Raises Exception if the token refresh fails (also marks token expired).
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

        return await self._exchange_refresh_token(email, refresh_token)

    async def _exchange_refresh_token(self, email: str, refresh_token: str) -> str:
        """Exchange a refresh token for an access token via PKCE refresh grant."""
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
                data = response.json()
                access_token = data["access_token"]
                expires_in = data.get("expires_in", 3600)

                # Cache
                self._cache[email] = (access_token, time.monotonic() + expires_in)

                # Persist rotated refresh token if present
                new_rt = data.get("refresh_token")
                if new_rt and new_rt != refresh_token:
                    enc = encrypt_token(new_rt, self._fernet_key)
                    await self._state.update_token(email, enc)
                    logger.debug("token_rotated", email=email)

                return access_token

            else:
                # Refresh failed — mark token expired
                await self._state.mark_token_expired(email)
                body = response.text
                logger.error(
                    "token_refresh_failed",
                    email=email,
                    status=response.status_code,
                    body=body[:200],
                )
                raise Exception(
                    f"Token refresh failed for {email}: HTTP {response.status_code}"
                )