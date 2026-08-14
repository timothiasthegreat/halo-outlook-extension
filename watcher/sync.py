"""Core sync engine — polls Graph API and pushes new messages to Halo tickets.

Supports multi-mailbox: loads watched mailboxes from state.db, creates a
per-user HaloClient for each mailbox, and routes conversations to the
correct user's token for accurate action attribution.

Usage:
    engine = SyncEngine(config, token_manager, graph, state)
    await engine.sync_once()
"""

from __future__ import annotations

import asyncio
from typing import Any

import structlog

from watcher.config import Config
from watcher.crypto import decrypt_token, get_fernet
from watcher.graph_client import GraphClient
from watcher.halo_client import HaloClient
from watcher.state import StateStore
from watcher.token_manager import TokenManager

logger = structlog.get_logger()


class SyncEngine:
    """Orchestrates the sync loop: Graph → state → Halo.

    For each watched conversation:
    1. Fetch all messages from Graph via conversationId filter (per-user mailbox)
    2. Dedup against synced_messages table
    3. Determine direction (sent vs received) based on the watching mailbox's email
    4. Post action to Halo ticket with correct outcome_id using that user's token
    5. Download and attach any file attachments
    6. Mark messages as synced
    """

    def __init__(
        self,
        config: Config,
        token_manager: TokenManager,
        graph: GraphClient,
        state: StateStore,
    ):
        self._config = config
        self._token_manager = token_manager
        self._graph = graph
        self._state = state

    async def sync_once(self) -> dict[str, int]:
        """Run one complete sync cycle. Creates per-user HaloClient instances."""
        stats: dict[str, int] = {
            "conversations_checked": 0,
            "messages_new": 0,
            "messages_synced": 0,
            "errors": 0,
        }

        watched = await self._state.get_watched_conversations()
        logger.info("sync_start", watched_count=len(watched))

        for conv in watched:
            stats["conversations_checked"] += 1
            conversation_id = str(conv["conversation_id"])
            ticket_id = int(conv["ticket_id"])
            watched_by = conv.get("watched_by")

            if not watched_by:
                logger.warning(
                    "conversation_no_watcher",
                    conversation_id=conversation_id,
                )
                continue

            try:
                async with HaloClient(
                    self._token_manager,
                    str(watched_by),
                    self._config.halo.api_url,
                    custom_field_conv_id=self._config.halo.custom_field_conv_id,
                    default_ticket_type_id=self._config.halo.default_ticket_type_id,
                ) as halo:
                    new_count = await self._sync_conversation(
                        conversation_id, ticket_id, str(watched_by), halo
                    )
                    stats["messages_synced"] += new_count
            except Exception:
                logger.exception(
                    "sync_conversation_failed",
                    conversation_id=conversation_id,
                    ticket_id=ticket_id,
                    watched_by=watched_by,
                )
                stats["errors"] += 1

            await asyncio.sleep(0.5)

        await self._handle_stale()
        logger.info("sync_complete", **stats)
        return stats

    async def _sync_conversation(
        self,
        conversation_id: str,
        ticket_id: int,
        watched_by: str,
        halo: HaloClient,
    ) -> int:
        """Sync one conversation's new messages to its ticket."""
        # Fetch messages from the user's mailbox
        messages = await self._graph.get_messages_by_conversation(
            conversation_id, user_email=watched_by
        )
        logger.debug(
            "graph_messages_fetched",
            conversation_id=conversation_id,
            total=len(messages),
        )

        synced_count = 0
        for msg in messages:
            internet_message_id = msg.get("internetMessageId", "")
            if not internet_message_id:
                logger.debug("message_no_internet_id", conversation_id=conversation_id)
                continue

            if await self._state.is_message_synced(conversation_id, internet_message_id):
                continue

            # Determine direction: is the sender the watcher's email?
            from_address = (
                msg.get("from", {})
                .get("emailAddress", {})
                .get("address", "")
                .lower()
            )
            is_from_user = from_address == watched_by.lower()

            outcome_id = (
                self._config.halo.actions.email_sent
                if is_from_user
                else self._config.halo.actions.email_received
            )
            subject = msg.get("subject", "(no subject)")
            body_preview = msg.get("bodyPreview", "")
            body_html = msg.get("body", {}).get("content", f"<p>{body_preview}</p>")

            note = f"Subject: {subject}\nFrom: {from_address}\n\n{body_preview}"
            note_html = (
                "<p><strong>Subject:</strong> "
                f"{_escape_html(subject)}<br>"
                f"<strong>From:</strong> {_escape_html(from_address)}</p>"
                f"{body_html}"
            )

            direction_label = "outbound" if is_from_user else "inbound"

            try:
                await halo.create_action(
                    ticket_id=ticket_id,
                    outcome_id=outcome_id,
                    note=note,
                    note_html=note_html,
                    email_message_id=internet_message_id,
                )
            except Exception:
                logger.exception(
                    "action_create_failed",
                    ticket_id=ticket_id,
                    direction=direction_label,
                    internet_message_id=internet_message_id[:50],
                )
                raise

            if msg.get("hasAttachments", False):
                await self._sync_attachments(ticket_id, msg, halo)

            await self._state.mark_synced(conversation_id, internet_message_id)
            synced_count += 1

        if synced_count:
            await self._state.touch_sync(conversation_id)

        return synced_count

    async def _sync_attachments(
        self, ticket_id: int, msg: dict[str, Any], halo: HaloClient
    ) -> None:
        """Download file attachments and attach them to the Halo ticket."""
        graph_id = msg.get("id", "")
        if not graph_id:
            return

        try:
            attachments = await self._graph.get_message_attachments(graph_id)
        except Exception:
            logger.warning("attachment_list_failed", ticket_id=ticket_id, exc_info=True)
            return

        for att in attachments:
            odata_type = att.get("@odata.type", "")
            if not odata_type.endswith("#microsoft.graph.fileAttachment"):
                continue

            try:
                content = await self._graph.get_attachment_content(graph_id, att["id"])
                if not content:
                    continue

                max_size = 10 * 1024 * 1024
                if len(content) > max_size:
                    logger.warning("attachment_too_large", name=att.get("name"), size=len(content))
                    continue

                await halo.attach_to_ticket(
                    ticket_id=ticket_id,
                    filename=att.get("name", "attachment"),
                    content=content,
                    content_type=att.get("contentType", "application/octet-stream"),
                )
            except Exception:
                logger.warning("attachment_download_failed", attachment_name=att.get("name"), exc_info=True)

    async def _handle_stale(self) -> None:
        """Find and mark stale conversations."""
        stale_days = self._config.watcher.stale_conversation_days
        stale_convs = await self._state.find_stale_conversations(stale_days)
        for conv_id in stale_convs:
            await self._state.mark_stale(conv_id)
        if stale_convs:
            logger.info("stale_conversations_marked", count=len(stale_convs), stale_days=stale_days)


def _escape_html(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")