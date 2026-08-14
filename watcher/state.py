"""SQLite-backed state store for the Halo Outlook Watcher.

Tracks which conversations are mapped to Halo tickets and which messages
have already been synced. WAL mode for concurrent read/write safety.

Schema:
  conversations:
    conversation_id TEXT PRIMARY KEY
    ticket_id INTEGER NOT NULL
    created_at TEXT NOT NULL
    last_sync_at TEXT
    is_stale INTEGER DEFAULT 0

  synced_messages:
    id INTEGER PRIMARY KEY AUTOINCREMENT
    conversation_id TEXT NOT NULL
    internet_message_id TEXT NOT NULL
    synced_at TEXT NOT NULL
    UNIQUE(conversation_id, internet_message_id)

  watched_mailboxes:
    email TEXT PRIMARY KEY
    refresh_token_enc TEXT NOT NULL
    token_status TEXT NOT NULL DEFAULT 'active'
    registered_at TEXT NOT NULL
    last_token_refresh_at TEXT
"""

from __future__ import annotations

import asyncio
import sqlite3
from datetime import datetime, timedelta, timezone

import aiosqlite
import structlog

logger = structlog.get_logger()

SCHEMA = """
CREATE TABLE IF NOT EXISTS conversations (
    conversation_id TEXT PRIMARY KEY,
    ticket_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    last_sync_at TEXT,
    is_stale INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS synced_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    conversation_id TEXT NOT NULL,
    internet_message_id TEXT NOT NULL,
    synced_at TEXT NOT NULL,
    UNIQUE(conversation_id, internet_message_id)
);

CREATE TABLE IF NOT EXISTS watched_mailboxes (
    email TEXT PRIMARY KEY,
    refresh_token_enc TEXT NOT NULL,
    token_status TEXT NOT NULL DEFAULT 'active',
    registered_at TEXT NOT NULL,
    last_token_refresh_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_synced_conv
    ON synced_messages(conversation_id);

CREATE INDEX IF NOT EXISTS idx_messages_dedup
    ON synced_messages(conversation_id, internet_message_id);
"""


class StateStore:
    """SQLite state store for conversation → ticket mappings.

    Usage:
        async with StateStore("state.db") as store:
            await store.track_conversation(conv_id, ticket_id)
            synced = await store.is_message_synced(conv_id, msg_id)
    """

    def __init__(self, db_path: str = "state.db"):
        self._db_path = db_path
        self._conn: aiosqlite.Connection | None = None

    # ── lifecycle ──────────────────────────────────────────────

    async def __aenter__(self) -> "StateStore":
        self._conn = await aiosqlite.connect(self._db_path)
        self._conn.row_factory = aiosqlite.Row
        # WAL mode for concurrent access
        await self._conn.execute("PRAGMA journal_mode=WAL")
        await self._conn.executescript(SCHEMA)
        await self._conn.commit()
        logger.info("state_store_opened", path=self._db_path)
        return self

    async def __aexit__(self, *args: object) -> None:
        if self._conn:
            await self._conn.close()
            self._conn = None
            logger.info("state_store_closed")

    # ── conversation tracking ──────────────────────────────────

    async def track_conversation(
        self, conversation_id: str, ticket_id: int
    ) -> None:
        """Record a conversation → ticket mapping.

        If the conversation is already tracked, updates last_sync_at.
        """
        now = _utcnow()
        assert self._conn is not None
        await self._conn.execute(
            """INSERT INTO conversations (conversation_id, ticket_id, created_at, last_sync_at)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(conversation_id) DO UPDATE SET
                   ticket_id = excluded.ticket_id,
                   last_sync_at = excluded.last_sync_at,
                   is_stale = 0""",
            (conversation_id, ticket_id, now, now),
        )
        await self._conn.commit()
        logger.debug("conversation_tracked", conversation_id=conversation_id, ticket_id=ticket_id)

    async def get_ticket_id(self, conversation_id: str) -> int | None:
        """Get the ticket ID for a tracked conversation, or None."""
        assert self._conn is not None
        cursor = await self._conn.execute(
            "SELECT ticket_id FROM conversations WHERE conversation_id = ? AND is_stale = 0",
            (conversation_id,),
        )
        row = await cursor.fetchone()
        return row["ticket_id"] if row else None

    async def get_watched_conversations(self) -> list[dict[str, object]]:
        """Get all active (non-stale) conversation → ticket mappings."""
        assert self._conn is not None
        cursor = await self._conn.execute(
            "SELECT conversation_id, ticket_id, last_sync_at FROM conversations WHERE is_stale = 0"
        )
        rows = await cursor.fetchall()
        return [dict(row) for row in rows]

    async def touch_sync(self, conversation_id: str) -> None:
        """Update last_sync_at for a conversation."""
        assert self._conn is not None
        await self._conn.execute(
            "UPDATE conversations SET last_sync_at = ?, is_stale = 0 WHERE conversation_id = ?",
            (_utcnow(), conversation_id),
        )
        await self._conn.commit()

    async def mark_stale(self, conversation_id: str) -> None:
        """Mark a conversation as stale (inactive)."""
        assert self._conn is not None
        await self._conn.execute(
            "UPDATE conversations SET is_stale = 1 WHERE conversation_id = ?",
            (conversation_id,),
        )
        await self._conn.commit()
        logger.info("conversation_stale", conversation_id=conversation_id)

    async def unlink_conversation(self, conversation_id: str) -> None:
        """Remove a conversation → ticket mapping entirely."""
        assert self._conn is not None
        await self._conn.execute(
            "DELETE FROM conversations WHERE conversation_id = ?",
            (conversation_id,),
        )
        await self._conn.execute(
            "DELETE FROM synced_messages WHERE conversation_id = ?",
            (conversation_id,),
        )
        await self._conn.commit()
        logger.info("conversation_unlinked", conversation_id=conversation_id)

    # ── message dedup ──────────────────────────────────────────

    async def is_message_synced(
        self, conversation_id: str, internet_message_id: str
    ) -> bool:
        """Check if a message has already been synced to the ticket."""
        assert self._conn is not None
        cursor = await self._conn.execute(
            "SELECT 1 FROM synced_messages WHERE conversation_id = ? AND internet_message_id = ?",
            (conversation_id, internet_message_id),
        )
        row = await cursor.fetchone()
        return row is not None

    async def mark_synced(
        self, conversation_id: str, internet_message_id: str
    ) -> bool:
        """Record a message as synced. Returns False if already synced (race)."""
        assert self._conn is not None
        now = _utcnow()
        try:
            await self._conn.execute(
                """INSERT INTO synced_messages (conversation_id, internet_message_id, synced_at)
                   VALUES (?, ?, ?)""",
                (conversation_id, internet_message_id, now),
            )
            await self._conn.commit()
            logger.debug("message_synced", conversation_id=conversation_id, msg_id=internet_message_id[:50])
            return True
        except aiosqlite.IntegrityError:
            # UNIQUE constraint — already synced
            return False

    # ── maintenance ────────────────────────────────────────────

    async def prune_stale_synced(self, days: int = 90) -> int:
        """Remove synced message records older than N days (cleanup).

        Returns count of deleted rows.
        """
        assert self._conn is not None
        cutoff = _utcnow_days_ago(days)
        cursor = await self._conn.execute(
            "DELETE FROM synced_messages WHERE synced_at < ?", (cutoff,)
        )
        await self._conn.commit()
        count = cursor.rowcount
        if count:
            logger.info("pruned_stale_synced", count=count, days=days)
        return count

    async def find_stale_conversations(self, stale_days: int) -> list[str]:
        """Find conversations that haven't been synced in N days.

        Returns list of conversation_ids.
        """
        assert self._conn is not None
        cutoff = _utcnow_days_ago(stale_days)
        cursor = await self._conn.execute(
            "SELECT conversation_id FROM conversations WHERE last_sync_at < ? AND is_stale = 0",
            (cutoff,),
        )
        rows = await cursor.fetchall()
        return [row["conversation_id"] for row in rows]

    # ── watched mailboxes ──────────────────────────────────────

    async def register_mailbox(self, email: str, refresh_token_enc: str) -> None:
        """Register or update a watched mailbox with its encrypted refresh token.

        Upsert: if the email already exists, updates the token and resets status
        to active. Otherwise inserts a new row.
        """
        assert self._conn is not None
        now = _utcnow()
        await self._conn.execute(
            """INSERT INTO watched_mailboxes
               (email, refresh_token_enc, token_status, registered_at, last_token_refresh_at)
               VALUES (?, ?, 'active', ?, ?)
               ON CONFLICT(email) DO UPDATE SET
                   refresh_token_enc = excluded.refresh_token_enc,
                   token_status = 'active',
                   last_token_refresh_at = excluded.last_token_refresh_at""",
            (email, refresh_token_enc, now, now),
        )
        await self._conn.commit()
        logger.info("mailbox_registered", email=email)

    async def get_watched_mailboxes(self) -> list[dict[str, object]]:
        """Get all active (non-expired) watched mailboxes with their tokens."""
        assert self._conn is not None
        cursor = await self._conn.execute(
            """SELECT email, refresh_token_enc, token_status,
                      registered_at, last_token_refresh_at
               FROM watched_mailboxes
               WHERE token_status = 'active'
               ORDER BY email"""
        )
        rows = await cursor.fetchall()
        return [dict(row) for row in rows]

    async def get_mailbox(self, email: str) -> dict[str, object] | None:
        """Get a single mailbox row by email, or None."""
        assert self._conn is not None
        cursor = await self._conn.execute(
            """SELECT email, refresh_token_enc, token_status,
                      registered_at, last_token_refresh_at
               FROM watched_mailboxes
               WHERE email = ?""",
            (email,),
        )
        row = await cursor.fetchone()
        return dict(row) if row else None

    async def update_token(self, email: str, refresh_token_enc: str) -> None:
        """Update the refresh token for a mailbox and reset status to active."""
        assert self._conn is not None
        await self._conn.execute(
            """UPDATE watched_mailboxes
               SET refresh_token_enc = ?,
                   token_status = 'active',
                   last_token_refresh_at = ?
               WHERE email = ?""",
            (refresh_token_enc, _utcnow(), email),
        )
        await self._conn.commit()
        logger.info("mailbox_token_updated", email=email)

    async def mark_token_expired(self, email: str) -> None:
        """Mark a mailbox's token as expired (pauses syncing for this user)."""
        assert self._conn is not None
        await self._conn.execute(
            "UPDATE watched_mailboxes SET token_status = 'expired' WHERE email = ?",
            (email,),
        )
        await self._conn.commit()
        logger.info("mailbox_token_expired", email=email)

    async def unregister_mailbox(self, email: str) -> None:
        """Remove a mailbox entirely from the watch list."""
        assert self._conn is not None
        await self._conn.execute(
            "DELETE FROM watched_mailboxes WHERE email = ?",
            (email,),
        )
        await self._conn.commit()
        logger.info("mailbox_unregistered", email=email)


def _utcnow() -> str:
    """ISO 8601 UTC timestamp string."""
    return datetime.now(timezone.utc).isoformat()


def _utcnow_days_ago(days: int) -> str:
    """ISO 8601 UTC timestamp for N days ago."""
    dt = datetime.now(timezone.utc) - timedelta(days=days)
    return dt.isoformat()