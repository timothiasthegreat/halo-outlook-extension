"""Async HTTP client for HaloPSA REST API with per-user token support.

Uses a TokenManager to obtain per-user access tokens instead of
client-credentials. The constructor takes a TokenManager and user email;
token lifecycle is managed externally.

All ticket/action endpoints require JSON arrays (Halo quirk: POST body
must be [{...}], not {...}).

Usage:
    tm = TokenManager(halo_config_dict, state_store, fernet_key)
    async with HaloClient(tm, "user@example.com", api_url) as client:
        tickets = await client.list_tickets(active_only=True)
"""

from __future__ import annotations

import base64
import time
from typing import Any

import httpx
import structlog

logger = structlog.get_logger()


class HaloClient:
    """Async client for HaloPSA API with per-user OAuth2 token management.

    Token acquisition and refresh is delegated to TokenManager.
    Each instance is bound to a single user's mailbox for attribution.
    """

    def __init__(
        self,
        token_manager: "TokenManager",
        user_email: str,
        api_url: str,
        *,
        custom_field_conv_id: int = 285,
        default_ticket_type_id: int = 1,
    ):
        from watcher.token_manager import TokenManager  # noqa: F811

        self._token_manager: TokenManager = token_manager
        self._user_email = user_email
        self._api_url = api_url.rstrip("/")
        self._custom_field_conv_id = custom_field_conv_id
        self._default_ticket_type_id = default_ticket_type_id
        self._client: httpx.AsyncClient | None = None

    # ── lifecycle ──────────────────────────────────────────────

    async def __aenter__(self) -> "HaloClient":
        self._client = httpx.AsyncClient(
            base_url=self._api_url,
            timeout=httpx.Timeout(30.0),
            headers={"Accept": "application/json"},
        )
        return self

    async def __aexit__(self, *args: Any) -> None:
        if self._client:
            await self._client.aclose()
            self._client = None

    # ── request helpers ────────────────────────────────────────

    async def _request(
        self,
        method: str,
        path: str,
        *,
        json: Any = None,
        params: dict[str, Any] | None = None,
    ) -> httpx.Response:
        """Make an authenticated request with retry on 429/5xx."""
        assert self._client is not None

        max_retries = 3
        for attempt in range(max_retries):
            access_token = await self._token_manager.get_token_for(self._user_email)
            response = await self._client.request(
                method,
                path,
                json=json,
                params=params,
                headers={"Authorization": f"Bearer {access_token}"},
            )

            if response.status_code < 500 and response.status_code != 429:
                return response

            delay = 2**attempt
            logger.warning(
                "retrying_request",
                method=method,
                path=path,
                status=response.status_code,
                attempt=attempt + 1,
                delay=delay,
            )
            await _async_sleep(delay)

        return response

    # ── tickets ────────────────────────────────────────────────

    async def create_ticket(
        self,
        summary: str,
        details_html: str,
        *,
        tickettype_id: int | None = None,
        user_id: int | None = None,
        customfield_value: str | None = None,
    ) -> dict[str, Any]:
        """Create a new ticket."""
        payload: dict[str, Any] = {
            "summary": summary[:70],
            "details_html": details_html,
            "tickettype_id": tickettype_id or self._default_ticket_type_id,
        }
        if user_id:
            payload["user_id"] = user_id
        if customfield_value:
            payload["customfields"] = [
                {"id": self._custom_field_conv_id, "value": customfield_value}
            ]

        response = await self._request("POST", "/Tickets", json=[payload])
        response.raise_for_status()
        data = response.json()
        ticket = data[0] if isinstance(data, list) else data
        logger.info("ticket_created", ticket_id=ticket["id"])
        return ticket

    async def get_ticket(self, ticket_id: int) -> dict[str, Any]:
        """Get a single ticket by ID."""
        response = await self._request("GET", f"/Tickets/{ticket_id}")
        response.raise_for_status()
        return response.json()

    async def update_ticket_custom_field(
        self, ticket_id: int, value: str
    ) -> dict[str, Any]:
        """Update the conversationId custom field on a ticket."""
        payload = {
            "id": ticket_id,
            "customfields": [
                {"id": self._custom_field_conv_id, "value": value}
            ],
        }
        response = await self._request("POST", "/Tickets", json=[payload])
        response.raise_for_status()
        logger.info("ticket_custom_field_updated", ticket_id=ticket_id)
        return response.json()

    async def search_tickets(self, query: str) -> list[dict[str, Any]]:
        """Search tickets by reference number or customer name."""
        response = await self._request(
            "GET", "/Tickets", params={"search": query, "$top": "20"}
        )
        response.raise_for_status()
        data = response.json()
        return data.get("tickets", [])

    # ── actions ────────────────────────────────────────────────

    async def create_action(
        self,
        ticket_id: int,
        outcome_id: int,
        note: str,
        note_html: str,
        *,
        email_message_id: str | None = None,
        sendemail: bool = False,
        hiddenfromuser: bool = False,
    ) -> dict[str, Any]:
        """Create a ticket action (journal entry)."""
        payload: dict[str, Any] = {
            "ticket_id": ticket_id,
            "outcome_id": outcome_id,
            "note": note,
            "note_html": note_html,
            "hiddenfromuser": hiddenfromuser,
            "sendemail": sendemail,
        }
        if email_message_id:
            payload["email_message_id"] = email_message_id

        response = await self._request("POST", "/Actions", json=[payload])
        response.raise_for_status()
        data = response.json()
        action = data[0] if isinstance(data, list) else data
        logger.info(
            "action_created",
            ticket_id=ticket_id,
            action_id=action["id"],
            outcome_id=outcome_id,
        )
        return action

    async def list_actions(
        self, ticket_id: int, *, top: int = 100
    ) -> list[dict[str, Any]]:
        """List recent actions on a ticket."""
        response = await self._request(
            "GET", "/Actions",
            params={"ticket_id": ticket_id, "$top": str(top)},
        )
        response.raise_for_status()
        data = response.json()
        return data if isinstance(data, list) else data.get(
            "actions", data.get("value", [])
        )

    # ── attachments ────────────────────────────────────────────

    async def attach_to_ticket(
        self,
        ticket_id: int,
        filename: str,
        content: bytes,
        *,
        content_type: str = "application/octet-stream",
    ) -> dict[str, Any]:
        """Upload a file attachment to a ticket."""
        payload = [
            {
                "ticket_id": ticket_id,
                "filename": filename,
                "contenttype": content_type,
                "data_base64": base64.b64encode(content).decode("ascii"),
            }
        ]
        response = await self._request("POST", "/Attachments", json=payload)
        response.raise_for_status()
        data = response.json()
        logger.info(
            "attachment_uploaded",
            ticket_id=ticket_id,
            filename=filename,
            size=len(content),
        )
        return data[0] if isinstance(data, list) else data


async def _async_sleep(seconds: float) -> None:
    """Async sleep helper."""
    import asyncio

    await asyncio.sleep(seconds)