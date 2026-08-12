"""Core sync engine — polls Graph API and pushes new messages to Halo tickets."""

from __future__ import annotations

import asyncio
from typing import Any

import structlog

from watcher.config import Config
from watcher.graph_client import GraphClient
from watcher.halo_client import HaloClient
from watcher.state import StateStore

logger = structlog.get_logger()


class SyncEngine:
    """Orchestrates the sync loop: Graph → state → Halo.

    For each watched conversation:
    1. Fetch all messages from Graph via conversationId filter
    2. Dedup against synced_messages table
    3. Determine direction (sent vs received)
    4. Post action to Halo ticket with correct outcome_id
    5. Mark messages as synced

    Usage:
        async with StateStore(...) as state, HaloClient(...) as halo, GraphClient(...) as graph:
            engine = SyncEngine(config, halo, graph, state)
            await engine.sync_once()
    """

    def __init__(
        self,
        config: Config,
        halo: HaloClient,
        graph: GraphClient,
        state: StateStore,
    ):
        self._config = config
        self._halo = halo
        self._graph = graph
        self._state = state
        self._user_email = config.graph.user_email.lower()

    async def sync_once(self) -> dict[str, int]:
        """Run one complete sync cycle.

        Returns:
            Dict with sync statistics: {conversations_checked, messages_new, messages_synced, errors}.
        """
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

            try:
                new_count = await self._sync_conversation(conversation_id, ticket_id)
                stats["messages_synced"] += new_count
            except Exception:
                logger.exception(
                    "sync_conversation_failed",
                    conversation_id=conversation_id,
                    ticket_id=ticket_id,
                )
                stats["errors"] += 1

            # Rate-limit between conversations
            await asyncio.sleep(0.5)

        # Mark stale conversations
        await self._handle_stale()

        logger.info("sync_complete", **stats)
        return stats

    async def _sync_conversation(self, conversation_id: str, ticket_id: int) -> int:
        """Sync one conversation's new messages to its ticket.

        Returns count of newly synced messages.
        """
        # Fetch all messages in this conversation
        messages = await self._graph.get_messages_by_conversation(conversation_id)
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

            # Dedup: skip if already synced
            if await self._state.is_message_synced(conversation_id, internet_message_id):
                continue

            # Determine direction
            from_address = (
                msg.get("from", {})
                .get("emailAddress", {})
                .get("address", "")
                .lower()
            )
            is_from_user = from_address == self._user_email

            # Build action payload
            outcome_id = (
                self._config.halo.actions.email_sent
                if is_from_user
                else self._config.halo.actions.email_received
            )
            subject = msg.get("subject", "(no subject)")
            body_preview = msg.get("bodyPreview", "")
            body_html = msg.get("body", {}).get("content", f"<p>{body_preview}</p>")

            note = f"Subject: {subject}\nFrom: {from_address}\n\n{body_preview}"
            note_html = f"<p><strong>Subject:</strong> {subject}<br><strong>From:</strong> {from_address}</p>{body_html}"

            direction_label = "outbound" if is_from_user else "inbound"

            # Post action to Halo
            try:
                await self._halo.create_action(
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

            # Mark as synced
            await self._state.mark_synced(conversation_id, internet_message_id)
            synced_count += 1
            logger.debug(
                "message_synced_to_action",
                ticket_id=ticket_id,
                direction=direction_label,
                subject=subject[:60],
            )

        if synced_count:
            await self._state.touch_sync(conversation_id)
            logger.info(
                "conversation_synced",
                conversation_id=conversation_id,
                ticket_id=ticket_id,
                new_messages=synced_count,
            )

        return synced_count

    async def _handle_stale(self) -> None:
        """Find and mark stale conversations."""
        stale_days = self._config.watcher.stale_conversation_days
        stale_convs = await self._state.find_stale_conversations(stale_days)
        for conv_id in stale_convs:
            await self._state.mark_stale(conv_id)