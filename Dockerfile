# Halo Outlook Extension — Watcher Dockerfile
# Multi-stage build for a small production image.

FROM python:3.11-slim AS builder

WORKDIR /app

# Install build dependencies
COPY watcher/ watcher/
COPY config.example.yaml ./

# Build watcher only
RUN pip install --no-cache-dir watcher/

# ── production stage ──────────────────────────────────────────

FROM python:3.11-slim

WORKDIR /app

# Copy installed packages from builder
COPY --from=builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=builder /usr/local/bin/uvicorn /usr/local/bin/uvicorn
COPY --from=builder /app/watcher /app/watcher

# Create non-root user
RUN useradd --create-home --shell /bin/bash watcher && \
    chown -R watcher:watcher /app

USER watcher

# Health check
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD python -c "import sys; sys.exit(0)"

# Default: daemon mode (override with --once for cron)
ENTRYPOINT ["python", "-m", "watcher.watcher"]
CMD ["--config", "/app/config.yaml"]