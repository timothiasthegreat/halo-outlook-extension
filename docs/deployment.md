# Deploying the Halo Outlook Extension

The watcher and static server run as a **single Docker container** by default.
Bare-metal (Python + Node.js) is also supported.

> **Prerequisites:** `config.yaml` filled in and `setup_check.py` passing.

---

## Deployment Options

| Method | Best for | Complexity |
|---|---|---|
| [Docker](#1-docker-recommended) | One command, reproducible, always-on. | Low |
| [Windows Service](#2-windows-service) | Windows MSPs with existing Windows Server. | Medium |
| [Linux systemd](#3-linux-systemd) | Linux servers with systemd. | Medium |
| [Cron](#4-cron-lightweight) | Lightweight, no daemon. | Low |

---

## 1. Docker (Recommended)

One container runs both Express (static + API on port 3000) and the watcher daemon (health on port 8888, internal only).

```bash
# 1. Build the add-in (one-time)
cd add-in && npm install && npm run build && cd ..

# 2. Copy and edit config (only halo.client_id + graph.* are required)
cp config.example.yaml config.yaml

# 3. Start
docker compose up -d

# 4. Verify
curl http://localhost:3000/health        # Express
curl http://localhost:3000/api/config    # Tenant config
```

### Reverse Proxy (required for production)

The container serves HTTP only. **You must put a TLS-terminating reverse proxy in front for any deployment that is not purely local development.** Outlook enforces HTTPS for all add-in assets, and the bearer tokens flowing through `/api/register` are sensitive.

**Recommended: Caddy** (automatic Let's Encrypt, 10 lines of config).

Add to `docker-compose.yml`:

```yaml
  caddy:
    image: caddy:2-alpine
    restart: unless-stopped
    ports:
      - "443:443"
      - "80:80"    # for HTTP→HTTPS redirect + Let's Encrypt
    volumes:
      - ./Caddyfile:/etc/caddy/Caddyfile:ro
      - caddy_data:/data
```

Create `Caddyfile`:

```
your-domain.com {
    reverse_proxy halo-outlook:3000
}
```

Then remove the public port mapping from the `halo-outlook` service:

```yaml
  halo-outlook:
    # ports:          # ← REMOVE the public mapping
    #   - "3000:3000"
    expose:
      - "3000"        # ← internal-only, reachable by Caddy
```

**Nginx** also works. Either way, the reverse proxy:

- Terminates TLS (Outlook requires HTTPS for add-in assets)
- Protects bearer tokens in transit (plain HTTP leaks credentials)
- Enables rate limiting and access logging at the edge

For local development with `webpack serve`, TLS is handled by webpack-dev-server — no reverse proxy needed.

### Volumes

| Path | Purpose |
|---|---|
| `./config.yaml:/app/config.yaml:ro` | Configuration (read-only) |
| `watcher_state:/app/state` | Persistent `state.db` + `fernet.key` |

### Environment variables in Docker

The Docker entrypoint auto-generates `FERNET_KEY` on first run and persists it to `/app/state/fernet.key`. No manual env setup needed. See [Configuration](configuration.md) for all options.

### Logs

```bash
docker compose logs -f
docker compose logs --tail=50
```

### Updating

```bash
docker compose down
docker compose build --pull
docker compose up -d
```

---

## 2. Windows Service

Use **NSSM** (Non-Sucking Service Manager) to wrap both processes.

```powershell
# Install
cd C:\halo-outlook-extension\watcher
pip install -e .
cd ..\server
npm ci --production
cd ..

# Create service directory
mkdir C:\halo-outlook-watcher
copy config.yaml C:\halo-outlook-watcher\

# Install watcher service
nssm install HaloWatcher python.exe
nssm set HaloWatcher AppDirectory "C:\halo-outlook-watcher"
nssm set HaloWatcher AppParameters "-m watcher.watcher --config config.yaml"
nssm set HaloWatcher Start SERVICE_AUTO_START

# Start Express server (separate service or task)
nssm install HaloServer node.exe
nssm set HaloServer AppDirectory "C:\halo-outlook-extension\server"
nssm set HaloServer AppParameters "server.js --config C:\halo-outlook-watcher\config.yaml"
nssm set HaloServer Start SERVICE_AUTO_START
```

Set `FERNET_KEY` as a system environment variable or in the service config.

---

## 3. Linux systemd

Two service units: one for the watcher, one for Express. Or run both from one service with `&`.

### Install

```bash
cd /opt
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension
cd watcher && pip install -e . && cd ..
cd server && npm ci --production && cd ..
cd add-in && npm install && npm run build && cd ..

# Fernet key
export FERNET_KEY=$(python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
echo "FERNET_KEY=$FERNET_KEY" > /etc/halo-outlook-watcher/env
```

### systemd unit

```ini
[Unit]
Description=Halo Outlook Extension
After=network-online.target

[Service]
Type=simple
User=halo-watcher
WorkingDirectory=/opt/halo-outlook-extension
EnvironmentFile=/etc/halo-outlook-watcher/env
ExecStart=/opt/halo-outlook-extension/docker-entrypoint.sh
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Or split into two units: `watcher.watcher` for Python, `node server/server.js` for Express.

### Logs

```bash
sudo journalctl -u halo-outlook-extension -f
```

---

## 4. Cron (Lightweight)

Run `--once` mode via cron:

```cron
*/2 * * * * /usr/bin/flock -n /tmp/halo-watcher.lock /usr/bin/python3 -m watcher.watcher --once --config /etc/halo-outlook-watcher/config.yaml
```

Express must be kept running separately (systemd or a process manager).

---

## Verifying

```bash
curl http://localhost:3000/health        # {"status":"ok"}
curl http://localhost:3000/api/config    # tenant config
```

The watcher health endpoint (port 8888) is internal-only — not exposed to the host in Docker by default.

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| `FERNET_KEY environment variable is required` | Key not set | Run Docker (auto-generates) or set manually |
| `halo.instance_url must start with https://` | Config URL has `http://` | Change to `https://` |
| `Graph connection failed` | Wrong credentials | Verify app registration has `Mail.Read` + admin consent |
| Health endpoint unreachable | Wrong port or firewall | Express: 3000 only |
| Messages not journaling | No users registered | Open the add-in, click "Enable Watching" |
| Add-in fails to load in Outlook | Missing TLS | Add a reverse proxy with HTTPS cert (see Reverse Proxy section above) |