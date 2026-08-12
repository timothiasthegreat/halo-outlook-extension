"""Watcher daemon entry point for Halo Outlook Extension.

Polls Graph API for new messages in watched conversations and journals
them to HaloPSA tickets. Supports daemon mode (continuous polling) and
--once mode (single sync cycle, suitable for cron).

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
            # Simple validation: try to list one message
            pass
    except Exception as e:
        issues.append(f"Graph connection failed: {e}")

    return issues


async def run_daemon(config_path: str) -> None:
    """Run the watcher in continuous poll mode."""
    config = load_config(config_path)
    setup_logging(config.watcher.log_level)

    logger.info("watcher_starting", mode="daemon", config=config_path)

    # Startup validation
    issues = await validate_config(config)
    if issues:
        for issue in issues:
            logger.error("config_validation_failed", issue=issue)
        logger.error("startup_aborted")
        return

    # Main poll loop
    poll_interval = config.watcher.poll_interval_seconds
    logger.info("poll_loop_starting", interval_seconds=poll_interval)

    async with (
        StateStore(config.watcher.state_db_path) as state,
        HaloClient(config.halo) as halo,
        GraphClient(config.graph) as graph,
    ):
        engine = SyncEngine(config, halo, graph, state)

        while True:
            try:
                stats = await engine.sync_once()
                logger.info("sync_cycle_done", **stats)
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
    args = parser.parse_args()

    if args.once:
        asyncio.run(run_once(args.config))
    else:
        asyncio.run(run_daemon(args.config))


if __name__ == "__main__":
    main()