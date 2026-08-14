#!/bin/bash
# Docker entrypoint — starts Express (background) and watcher (foreground).
# Generates FERNET_KEY on first run if not set and persists to volume.

set -e

# ── Fernet key management ──────────────────────────────────────
FERNET_KEY_FILE="${FERNET_KEY_FILE:-/app/state/fernet.key}"

if [ -z "${FERNET_KEY:-}" ]; then
  if [ -f "$FERNET_KEY_FILE" ]; then
    export FERNET_KEY
    FERNET_KEY="$(cat "$FERNET_KEY_FILE")"
  else
    # Generate a new Fernet key
    FERNET_KEY="$(python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")"
    echo "$FERNET_KEY" > "$FERNET_KEY_FILE"
    echo "Generated new Fernet key at $FERNET_KEY_FILE"
  fi
fi
export FERNET_KEY

# ── Start Express server in background ─────────────────────────
EXPRESS_PORT="${EXPRESS_PORT:-3000}"

node /app/server/server.js --port "$EXPRESS_PORT" --config /app/config.yaml &
EXPRESS_PID=$!
echo "Express server started (PID $EXPRESS_PID) on port $EXPRESS_PORT"

# ── Start watcher in foreground ────────────────────────────────
HEALTH_PORT="${HEALTH_PORT:-8888}"

exec python -m watcher.watcher --config /app/config.yaml --health-port "$HEALTH_PORT"