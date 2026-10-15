"""Watcher daemon entry point for Halo Outlook Extension.

Polls Graph API for new messages in watched conversations and journals
them to HaloPSA tickets. Supports daemon mode (continuous polling) and
--once mode (single sync cycle, suitable for cron).

Uses per-user PKCE tokens via TokenManager — no client credentials needed
for Halo. The watcher loads watched mailboxes from state.db and uses each
user's own refresh token for authenticating API calls.

Usage:
    python -m watcher.watcher                    # daemon mode
    python -m watcher.watcher --once             # single sync cycle
    python -m watcher.watcher --config config.yaml  # custom config path
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import sys

import structlog

from watcher.config import load_config
from watcher.crypto import get_fernet
from watcher.graph_client import GraphClient
from watcher.state import StateStore
from watcher.sync import SyncEngine
from watcher.token_manager import TokenManager

logger = structlog.get_logger()

HEALTH_PORT = 8888


def setup_logging(log_level: str = "INFO") -> None:
    """Configure structured JSON logging to stdout."""
    level = getattr(logging, log_level.upper(), logging.INFO)
    structlog.configure(
        processors=[
            structlog.stdlib.add_log_level,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.dev.ConsoleRenderer(),
        ],
        wrapper_class=structlog.stdlib.BoundLogger,
        context_class=dict,
        logger_factory=structlog.stdlib.LoggerFactory(),
        cache_logger_on_first_use=True,
    )
    logging.basicConfig(format="%(message)s", stream=sys.stdout, level=level)


async def validate_config(config) -> list[str]:
    """Validate configuration on startup."""
    issues: list[str] = []
    try:
        async with GraphClient(config.graph) as graph:
            pass
    except Exception as e:
        issues.append(f"Graph connection failed: {e}")
    return issues


def _build_health_app(shared_state: dict):
    """Build a tiny FastAPI app for health + state.db write-through.

    The Express server proxies /api/conversations and /api/register to
    this app so that all SQLite writes go through the single aiosqlite
    connection the watcher holds.  This avoids the sql.js ↔ aiosqlite
    WAL corruption that happens when two processes write the same file.
    """
    import httpx
    from fastapi import FastAPI, HTTPException
    from fastapi.responses import JSONResponse

    from watcher.token_manager import (
        RegistrationMismatchError,
        TokenRefreshError,
        TokenRevokedError,
    )

    app = FastAPI(title="Halo Outlook Watcher — API")

    @app.get("/health")
    async def health():
        return JSONResponse(content={
            "status": "ok",
            "conversations": shared_state.get("conversations", 0),
            "mailboxes": shared_state.get("mailboxes", 0),
            "last_sync_at": shared_state.get("last_sync_at"),
        })

    # ── state.db write-through endpoints ─────────────────────

    @app.post("/track-conversation")
    async def track_conversation(req: "dict"):
        """Write a conversation→ticket mapping to the watcher's state.db."""
        state = shared_state.get("_state_store")
        if state is None:
            raise HTTPException(503, "State store not initialized")
        conv_id = req.get("conversationId")
        ticket_id = req.get("ticketId")
        watched_by = req.get("watchedBy")
        if not conv_id or not ticket_id:
            raise HTTPException(400, "conversationId and ticketId required")
        await state.track_conversation(conv_id, int(ticket_id), watched_by)
        return {"status": "ok", "conversationId": conv_id, "ticketId": ticket_id}

    @app.get("/conversation")
    async def get_conversation(conversation_id: str, watched_by: str):
        """Look up a conversation→ticket mapping for the mailbox watching it.

        404 when the conversation isn't tracked for watched_by — a mapping
        another mailbox holds isn't reported.
        """
        state = shared_state.get("_state_store")
        if state is None:
            raise HTTPException(503, "State store not initialized")
        row = await state.get_conversation(conversation_id)
        if row is None or str(row.get("watched_by") or "").lower() != watched_by.strip().lower():
            raise HTTPException(404, "Conversation is not tracked")
        return {
            "conversationId": row["conversation_id"],
            "ticketId": row["ticket_id"],
            "lastSyncAt": row["last_sync_at"],
            "isStale": bool(row["is_stale"]),
        }

    @app.post("/untrack-conversation")
    async def untrack_conversation(req: "dict"):
        """Stop watching a conversation (the add-in's Unlink).

        Only the mailbox watching the conversation can untrack it.
        """
        state = shared_state.get("_state_store")
        if state is None:
            raise HTTPException(503, "State store not initialized")
        conv_id = req.get("conversationId")
        watched_by = str(req.get("watchedBy") or "").strip().lower()
        if not conv_id or not watched_by:
            raise HTTPException(400, "conversationId and watchedBy required")
        row = await state.get_conversation(conv_id)
        if row is None:
            raise HTTPException(404, "Conversation is not tracked")
        if str(row.get("watched_by") or "").lower() != watched_by:
            raise HTTPException(403, "Conversation is watched from another mailbox")
        await state.unlink_conversation(conv_id)
        return {"status": "ok", "conversationId": conv_id}

    @app.post("/register-mailbox")
    async def register_mailbox(req: "dict"):
        """Register a watched mailbox in the watcher's state.db.

        The refresh token must belong to the Halo agent with this email —
        TokenManager.register_verified() checks that with Halo before
        anything is stored, so a caller cannot overwrite another user's token.
        """
        token_manager = shared_state.get("_token_manager")
        if token_manager is None:
            raise HTTPException(503, "Token manager not initialized")

        email = req.get("email", "").strip().lower()
        refresh_token = req.get("refresh_token", "")

        if not email or "@" not in email:
            raise HTTPException(400, "Valid email required")
        if not refresh_token:
            raise HTTPException(400, "refresh_token required")

        try:
            await token_manager.register_verified(email, refresh_token)
        except RegistrationMismatchError as e:
            raise HTTPException(403, str(e))
        except TokenRevokedError:
            raise HTTPException(400, "Halo rejected the refresh token")
        except (TokenRefreshError, httpx.HTTPError) as e:
            raise HTTPException(502, f"Could not verify token with Halo: {e}")

        return {"status": "ok", "email": email, "watching": [email]}

    @app.get("/mailbox-status")
    async def mailbox_status(email: str):
        """Report whether a mailbox is registered and its token is usable."""
        state = shared_state.get("_state_store")
        if state is None:
            raise HTTPException(503, "State store not initialized")
        row = await state.get_mailbox(email.strip().lower())
        return {
            "email": email,
            "registered": row is not None,
            "token_status": row["token_status"] if row else None,
        }

    return app


async def _run_health_server(shared_state: dict, port: int) -> None:
    """Start the health-check server in a background thread."""
    import threading
    import uvicorn

    app = _build_health_app(shared_state)
    server = uvicorn.Server(uvicorn.Config(app, host="0.0.0.0", port=port, log_level="warning"))
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    await asyncio.sleep(0.1)


async def run_daemon(config_path: str, health_port: int = HEALTH_PORT) -> None:
    """Run the watcher in continuous poll mode."""
    config = load_config(config_path)
    setup_logging(config.watcher.log_level)

    logger.info("watcher_starting", mode="daemon", config=config_path, health_port=health_port)

    issues = await validate_config(config)
    if issues:
        for issue in issues:
            logger.error("config_validation_failed", issue=issue)
        logger.error("startup_aborted")
        return

    shared_state: dict = {"conversations": 0, "last_sync_at": None, "mailboxes": 0}
    poll_interval = config.watcher.poll_interval_seconds
    logger.info("poll_loop_starting", interval_seconds=poll_interval)

    async with (
        StateStore(config.watcher.state_db_path) as state,
        GraphClient(config.graph) as graph,
    ):
        shared_state["_state_store"] = state

        # TokenManager needs the Fernet key as a base64 string so it can
        # decrypt tokens that were encrypted by get_fernet() (which also
        # reads FERNET_KEY from the environment).
        fernet_key_str = os.environ["FERNET_KEY"]
        halo_config_dict = config.halo.model_dump()
        token_manager = TokenManager(halo_config_dict, state, fernet_key_str)
        shared_state["_token_manager"] = token_manager
        engine = SyncEngine(config, token_manager, graph, state)

        await _run_health_server(shared_state, health_port)
        logger.info("health_server_started", port=health_port)

        while True:
            try:
                stats = await engine.sync_once()
                logger.info("sync_cycle_done", **stats)
                watched = await state.get_watched_conversations()
                mailboxes = await state.get_watched_mailboxes()
                shared_state["conversations"] = len(watched)
                shared_state["mailboxes"] = len(mailboxes)
                shared_state["last_sync_at"] = _utcnow_iso()
            except Exception:
                logger.exception("sync_cycle_failed")
            logger.debug("sleeping", seconds=poll_interval)
            await asyncio.sleep(poll_interval)


async def run_once(config_path: str) -> None:
    """Run a single sync cycle and exit."""
    config = load_config(config_path)
    setup_logging(config.watcher.log_level)

    logger.info("watcher_starting", mode="once", config=config_path)

    async with (
        StateStore(config.watcher.state_db_path) as state,
        GraphClient(config.graph) as graph,
    ):
        fernet_key_str = os.environ["FERNET_KEY"]
        halo_config_dict = config.halo.model_dump()
        token_manager = TokenManager(halo_config_dict, state, fernet_key_str)
        engine = SyncEngine(config, token_manager, graph, state)
        stats = await engine.sync_once()
        logger.info("sync_once_done", **stats)


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(description="Halo Outlook Extension — Watcher Daemon")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--once", action="store_true", help="Run a single sync cycle and exit")
    parser.add_argument("--health-port", type=int, default=HEALTH_PORT, help="Health-check port")
    args = parser.parse_args()

    if args.once:
        asyncio.run(run_once(args.config))
    else:
        asyncio.run(run_daemon(args.config, health_port=args.health_port))


def _utcnow_iso() -> str:
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).isoformat()


if __name__ == "__main__":
    main()