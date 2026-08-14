"""Async HTTP client for Microsoft Graph API."""

from __future__ import annotations

import time
from typing import Any

import httpx
import structlog

from watcher.config import GraphConfig

logger = structlog.get_logger()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"


class GraphClient:
    """Async client for Microsoft Graph API with token management.

    Uses OAuth2 client credentials grant (application permission Mail.Read).
    Supports per-user mailbox queries via the user_email parameter.

    Usage:
        async with GraphClient(config) as client:
            messages = await client.get_messages_by_conversation(conv_id)
    """

    def __init__(self, config: GraphConfig):
        self._config = config
        self._client: httpx.AsyncClient | None = None
        self._access_token: str | None = None
        self._token_expiry: float = 0.0

    # ── lifecycle ──────────────────────────────────────────────

    async def __aenter__(self) -> "GraphClient":
        self._client = httpx.AsyncClient(
            base_url=GRAPH_BASE_URL,
            timeout=httpx.Timeout(30.0),
            headers={"Accept": "application/json"},
        )
        await self._ensure_token()
        return self

    async def __aexit__(self, *args: Any) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None

    # ── auth ───────────────────────────────────────────────────

    async def _ensure_token(self) -> None:
        """Acquire a new access token if needed."""
        now = time.monotonic()
        if self._access_token and now < self._token_expiry - 60:
            return
        await self._acquire_token()

    async def _acquire_token(self) -> None:
        """Acquire token via client credentials grant."""
        assert self._client is not None
        token_url = TOKEN_URL.format(tenant_id=self._config.tenant_id)
        response = await self._client.post(
            token_url,
            data={
                "grant_type": "client_credentials",
                "client_id": self._config.client_id,
                "client_secret": self._config.client_secret,
                "scope": "https://graph.microsoft.com/.default",
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        response.raise_for_status()
        data = response.json()
        self._access_token = data["access_token"]
        self._token_expiry = time.monotonic() + data.get("expires_in", 3600)
        logger.info("graph_token_acquired")

    # ── request helpers ────────────────────────────────────────

    async def _request(
        self, path: str, *, params: dict[str, Any] | None = None
    ) -> httpx.Response:
        """Make an authenticated request with retry on 429/5xx."""
        assert self._client is not None

        max_retries = 3
        for attempt in range(max_retries):
            await self._ensure_token()
            response = await self._client.get(
                path,
                params=params,
                headers={"Authorization": f"Bearer {self._access_token}"},
            )

            if response.status_code < 500 and response.status_code != 429:
                return response

            delay = 2**attempt
            logger.warning(
                "graph_retry",
                path=path,
                status=response.status_code,
                attempt=attempt + 1,
                delay=delay,
            )
            await _async_sleep(delay)

        return response

    async def _paginate(
        self, path: str, params: dict[str, Any] | None = None
    ) -> list[dict[str, Any]]:
        """Fetch all pages of a paginated Graph API response."""
        params = dict(params or {})
        params.setdefault("$top", "50")

        all_items: list[dict[str, Any]] = []
        next_url: str | None = path

        while next_url:
            if next_url.startswith("https://"):
                from urllib.parse import urlparse

                parsed = urlparse(next_url)
                next_url = parsed.path + ("?" + parsed.query if parsed.query else "")
                page_params = None
            else:
                page_params = params

            response = await self._request(next_url, params=page_params)
            if response.status_code != 200:
                logger.error("graph_paginate_error", status=response.status_code)
                break

            data = response.json()
            if "error" in data:
                logger.error("graph_api_error", error=data["error"])
                break

            items = data.get("value", [])
            all_items.extend(items)
            next_url = data.get("@odata.nextLink")
            params = None

            if not items:
                break

        return all_items

    # ── messages ───────────────────────────────────────────────

    async def get_messages_by_conversation(
        self, conversation_id: str, *, user_email: str | None = None
    ) -> list[dict[str, Any]]:
        """Fetch all messages in a given conversation.

        If user_email is provided, queries that specific mailbox.
        Otherwise uses the configured user_email from GraphConfig.
        """
        email = user_email or self._config.user_email
        user_path = f"/users/{email}/messages"
        params = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$select": (
                "subject,from,sentDateTime,internetMessageId,"
                "body,hasAttachments,id"
            ),
            "$top": "50",
        }
        return await self._paginate(user_path, params)

    async def get_messages_since(self, since: str) -> list[dict[str, Any]]:
        """Fetch messages since a given timestamp."""
        user_path = f"/users/{self._config.user_email}/messages"
        params = {
            "$filter": f"receivedDateTime ge {since}",
            "$select": (
                "subject,from,sentDateTime,internetMessageId,"
                "conversationId,hasAttachments,id"
            ),
            "$top": "50",
            "$orderby": "receivedDateTime desc",
        }
        return await self._paginate(user_path, params)

    async def get_message(self, message_id: str) -> dict[str, Any]:
        """Get a single message by ID."""
        response = await self._request(
            f"/users/{self._config.user_email}/messages/{message_id}"
        )
        response.raise_for_status()
        return response.json()

    # ── attachments ────────────────────────────────────────────

    async def get_message_attachments(
        self, message_id: str
    ) -> list[dict[str, Any]]:
        """Fetch attachment metadata for a message."""
        user_path = (
            f"/users/{self._config.user_email}"
            f"/messages/{message_id}/attachments"
        )
        response = await self._request(user_path)
        response.raise_for_status()
        data = response.json()
        return data.get("value", [])

    async def get_attachment_content(
        self, message_id: str, attachment_id: str
    ) -> bytes | None:
        """Download the raw bytes of a file attachment."""
        user_path = (
            f"/users/{self._config.user_email}/messages/{message_id}"
            f"/attachments/{attachment_id}/$value"
        )
        response = await self._request(user_path)
        if response.status_code == 200:
            return response.content
        return None


async def _async_sleep(seconds: float) -> None:
    """Async sleep helper."""
    import asyncio

    await asyncio.sleep(seconds)