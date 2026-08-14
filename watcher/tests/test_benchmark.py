"""Performance benchmark for SyncEngine with simulated load."""

import asyncio
import tempfile
import time
from unittest.mock import AsyncMock

from watcher.config import Config
from watcher.sync import SyncEngine
from watcher.state import StateStore


def make_config() -> Config:
    return Config.model_validate(
        {
            "halo": {
                "instance_url": "https://test.halopsa.com",
                "client_id": "test-client",
                "client_secret": "test-secret",
            },
            "graph": {
                "tenant_id": "test-tenant",
                "client_id": "test-graph-client",
                "client_secret": "test-graph-secret",
                "user_email": "user@example.com",
            },
        }
    )


async def run_benchmark(num_conversations: int, messages_per_conv: int = 3):
    """Run a sync cycle with simulated conversations and measure performance."""
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as state:
        config = make_config()
        halo = AsyncMock()
        graph = AsyncMock()

        # Build mock Graph response — all conversations return the same messages
        def _build_messages(conv_id: str):
            return [
                {
                    "subject": f"Test message {i}",
                    "from": {"emailAddress": {"address": "customer@example.com"}},
                    "sentDateTime": "2026-08-12T12:00:00Z",
                    "internetMessageId": f"<{conv_id}-msg-{i}@test.com>",
                    "hasAttachments": False,
                    "body": {"content": f"<p>Body {i}</p>"},
                }
                for i in range(messages_per_conv)
            ]

        # Track many conversations
        for i in range(num_conversations):
            conv_id = f"conv-bench-{i:06d}"
            await state.track_conversation(conv_id, 1000 + i)

        graph.get_messages_by_conversation = AsyncMock(
            return_value=_build_messages("conv")
        )
        halo.create_action = AsyncMock(return_value={"id": 1})
        halo.attach_to_ticket = AsyncMock()

        engine = SyncEngine(config, halo, graph, state)

        start = time.perf_counter()
        stats = await engine.sync_once()
        elapsed = time.perf_counter() - start

        assert stats["messages_synced"] == num_conversations * messages_per_conv
        assert stats["conversations_checked"] == num_conversations
        assert stats["errors"] == 0

        return elapsed


def test_benchmark_small():
    """Baseline: 10 conversations, 3 messages each."""
    elapsed = asyncio.run(run_benchmark(10, 3))
    # Should be well under 6 seconds (10 × 0.5s sleep between convs = 5s alone)
    assert elapsed < 7.0, f"Small benchmark took {elapsed:.2f}s (expected < 7s)"


def test_benchmark_medium():
    """Medium: 50 conversations, 3 messages each."""
    elapsed = asyncio.run(run_benchmark(50, 3))
    # 50 × 0.5s = 25s minimum due to rate-limiting sleep
    assert elapsed < 30.0, f"Medium benchmark took {elapsed:.2f}s (expected < 30s)"


def test_benchmark_large():
    """Large: 100 conversations, 3 messages each."""
    elapsed = asyncio.run(run_benchmark(100, 3))
    # 100 × 0.5s = 50s minimum
    assert elapsed < 55.0, f"Large benchmark took {elapsed:.2f}s (expected < 55s)"