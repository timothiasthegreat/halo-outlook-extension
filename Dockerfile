# Halo Outlook Extension — Combined Container
# Multi-stage: Node.js (Express) + Python (watcher) in one image.
#
# The container runs:
#   - Express.js on port 3000 (static add-in + /api/* endpoints)
#   - Watcher daemon with health-check on port 8888 (internal)
#
# Volumes:
#   /app/state  — persistent state.db + fernet.key
#   /app/config.yaml — configuration (read-only, bind-mounted)
#
# Environment:
#   FERNET_KEY        — encryption key (auto-generated if missing, persisted to volume)
#   FERNET_KEY_FILE   — path to fernet key file (default: /app/state/fernet.key)
#   EXPRESS_PORT      — Express listen port (default: 3000)
#   HEALTH_PORT       — watcher health-check port (default: 8888)
#   PYTHONUNBUFFERED  — set to 1 for unbuffered log output

# ── Node.js builder ────────────────────────────────────────────
FROM node:18-slim AS node-builder

WORKDIR /app/server
COPY server/package.json server/package-lock.json* ./
RUN npm ci --production

# ── Python builder ─────────────────────────────────────────────
FROM python:3.11-slim AS python-builder

WORKDIR /app
COPY watcher/ watcher/
RUN pip install --no-cache-dir watcher/

# ── Production image ───────────────────────────────────────────
FROM python:3.11-slim

# Install Node.js 18 for the Express server
RUN apt-get update && \
    apt-get install -y --no-install-recommends curl gnupg && \
    curl -fsSL https://deb.nodesource.com/setup_18.x | bash - && \
    apt-get install -y --no-install-recommends nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy Python packages from builder
COPY --from=python-builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=python-builder /usr/local/bin/uvicorn /usr/local/bin/uvicorn
COPY --from=python-builder /app/watcher /app/watcher

# Copy Node.js server
COPY --from=node-builder /app/server/node_modules /app/server/node_modules
COPY server/ /app/server/

# Copy add-in static files
COPY add-in/dist/ /app/add-in/dist/

# Copy entrypoint
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint
RUN chmod +x /usr/local/bin/docker-entrypoint

# Create state directory for volume mount
RUN mkdir -p /app/state

# Create non-root user
RUN useradd --create-home --shell /bin/bash watcher && \
    chown -R watcher:watcher /app

USER watcher

# Health check — probes Express /health endpoint
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD node -e "const http=require('http');http.get('http://localhost:3000/health',r=>{process.exit(r.statusCode===200?0:1)})"

# Expose Express port (public) and watcher health port (internal)
EXPOSE 3000 8888

ENTRYPOINT ["docker-entrypoint"]