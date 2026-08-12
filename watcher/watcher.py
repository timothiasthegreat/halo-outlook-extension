"""Watcher daemon entry point for Halo Outlook Extension.

Polls Graph API for new messages in watched conversations and journals
them to HaloPSA tickets. Supports daemon mode (continuous polling) and
--once mode (single sync cycle, suitable for cron).

In daemon mode, also starts a FastAPI health-check HTTP server on
localhost:8888 (configurable) so Docker / orchestrators can verify
the watcher is alive without scraping logs.

Usage:
    python -m watcher.watcher                    # daemon mode
    python -m watcher.watcher --once             # single sync cycle
    python -m watcher.watcher --config config.yaml  # custom config path
"""

from __future__ import annotations

import argparse
import asyncio
import signal
import sys
from pathlib import Path

import structlog

from watcher.config import HaloActionsConfig, load_config
from watcher.graph_client import GraphClient
from watcher.halo_client import HaloClient
from watcher.state import StateStore
from watcher.sync import SyncEngine

logger = structlog.get_logger()

# Default health-check port
HEALTH_PORT = 8888


def setup_logging(log_level: str = "INFO") -> None:
    """Configure structured JSON logging to stdout."""
    structlog.configure(
        processors=[
            structlog.stdlib.filter_by_level,
            structlog.stdlib.add_log_level,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.dev.ConsoleRenderer(),
        ],
        wrapper_class=structlog.stdlib.BoundLogger,
        context_class=dict,
        logger_factory=structlog.PrintLoggerFactory(),
        cache_logger_on_first_use=True,
    )


async def validate_config(config) -> list[str]:
    """Validate configuration on startup.

    Returns list of issues (empty = valid).
    """
    issues: list[str] = []

    # Check Halo reachable
    try:
        async with HaloClient(config.halo) as halo:
            pass  # Client init tests token acquisition
    except Exception as e:
        issues.append(f"Halo connection failed: {e}")

    # Check Graph reachable
    try:
        async with GraphClient(config.graph) as graph:
            pass
    except Exception as e:
        issues.append(f"Graph connection failed: {e}")

    return issues


# ── health-check server ────────────────────────────────────────


def _build_health_app(shared_state: dict) -> "fastapi.FastAPI":
    """Build a tiny FastAPI app that reads from the in-process shared state."""
    from fastapi import FastAPI
    from fastapi.responses import JSONResponse

    app = FastAPI(title="Halo Outlook Watcher — Health")

    @app.get("/health")
    async def health() -> JSONResponse:
        conversations = shared_state.get("conversations", 0)
        return JSONResponse(
            content={
                "status": "ok",
                "conversations": conversations,
                "last_sync_at": shared_state.get("last_sync_at"),
            }
        )

    return app


async def _run_health_server(shared_state: dict, port: int) -> None:
    """Start the health-check server in a background thread.

    uvicorn.Server.serve() blocks the event loop, so we delegate to
    a daemon thread that runs its own event loop. The thread dies
    when the process exits.
    """
    import threading

    import uvicorn

    app = _build_health_app(shared_state)
    config = uvicorn.Config(
        app,
        host="0.0.0.0",
        port=port,
        log_level="warning",
    )
    server = uvicorn.Server(config)

    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    # Give the socket a moment to bind before returning
    await asyncio.sleep(0.1)


# ── daemon / once-mode runners ─────────────────────────────────


async def run_daemon(config_path: str, health_port: int = HEALTH_PORT) -> None:
    """Run the watcher in continuous poll mode with health-check server."""
    config = load_config(config_path)
    setup_logging(config.watcher.log_level)

    logger.info(
        "watcher_starting", mode="daemon", config=config_path, health_port=health_port
    )

    # Startup validation
    issues = await validate_config(config)
    if issues:
        for issue in issues:
            logger.error("config_validation_failed", issue=issue)
        logger.error("startup_aborted")
        return

    # Shared state: the health endpoint reads from this dict
    shared_state: dict = {"conversations": 0, "last_sync_at": None}

    # Main poll loop
    poll_interval = config.watcher.poll_interval_seconds
    logger.info("poll_loop_starting", interval_seconds=poll_interval)

    async with (
        StateStore(config.watcher.state_db_path) as state,
        HaloClient(config.halo) as halo,
        GraphClient(config.graph) as graph,
    ):
        engine = SyncEngine(config, halo, graph, state)

        # Kick off the health server in a background thread
        await _run_health_server(shared_state, health_port)
        logger.info("health_server_started", port=health_port)

        while True:
            try:
                stats = await engine.sync_once()
                logger.info("sync_cycle_done", **stats)

                # Update shared state for health endpoint
                watched = await state.get_watched_conversations()
                shared_state["conversations"] = len(watched)
                shared_state["last_sync_at"] = _utcnow_iso()

            except Exception:
                logger.exception("sync_cycle_failed")
                # Continue polling — don't crash on transient failures

            logger.debug("sleeping", seconds=poll_interval)
            await asyncio.sleep(poll_interval)


async def run_once(config_path: str) -> None:
    """Run a single sync cycle and exit."""
    config = load_config(config_path)
    setup_logging(config.watcher.log_level)

    logger.info("watcher_starting", mode="once", config=config_path)

    async with (
        StateStore(config.watcher.state_db_path) as state,
        HaloClient(config.halo) as halo,
        GraphClient(config.graph) as graph,
    ):
        engine = SyncEngine(config, halo, graph, state)
        stats = await engine.sync_once()
        logger.info("sync_once_done", **stats)


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Halo Outlook Extension — Watcher Daemon"
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to config.yaml (default: config.yaml)",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run a single sync cycle and exit (suitable for cron)",
    )
    parser.add_argument(
        "--health-port",
        type=int,
        default=HEALTH_PORT,
        help=f"Port for health-check HTTP server (default: {HEALTH_PORT}, daemon mode only)",
    )
    args = parser.parse_args()

    if args.once:
        asyncio.run(run_once(args.config))
    else:
        asyncio.run(run_daemon(args.config, health_port=args.health_port))


def _utcnow_iso() -> str:
    """ISO 8601 UTC timestamp string."""
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).isoformat()


if __name__ == "__main__":
    main()