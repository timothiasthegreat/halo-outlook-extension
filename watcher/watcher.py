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
    """Build a tiny FastAPI app for health checks."""
    from fastapi import FastAPI
    from fastapi.responses import JSONResponse

    app = FastAPI(title="Halo Outlook Watcher — Health")

    @app.get("/health")
    async def health():
        return JSONResponse(content={
            "status": "ok",
            "conversations": shared_state.get("conversations", 0),
            "mailboxes": shared_state.get("mailboxes", 0),
            "last_sync_at": shared_state.get("last_sync_at"),
        })

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
        fernet_k = get_fernet()
        raw_key = fernet_k._signing_key + fernet_k._encryption_key
        halo_config_dict = config.halo.model_dump()
        token_manager = TokenManager(halo_config_dict, state, raw_key)
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
        fernet_k = get_fernet()
        raw_key = fernet_k._signing_key + fernet_k._encryption_key
        halo_config_dict = config.halo.model_dump()
        token_manager = TokenManager(halo_config_dict, state, raw_key)
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