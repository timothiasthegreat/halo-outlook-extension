# Deploying the Halo Outlook Extension Watcher

The watcher is a background service that polls Microsoft Graph for new
messages in tracked conversations and journals them to HaloPSA tickets.
It runs the same Python code everywhere — only the process wrapper differs.

> **Prerequisites:** You've already [configured](configuration.md) your
> `config.yaml` and validated connectivity with `setup_check.py`.

---

## Deployment Options

| Method | Best for | Complexity |
|---|---|---|
| [Docker](#1-docker-recommended) | Anyone comfortable with Docker. One command. | Low |
| [Windows Service](#2-windows-service) | Windows MSPs using their existing Windows Server | Medium |
| [Linux systemd](#3-linux-systemd) | Linux servers with systemd | Medium |
| [Cron](#4-cron-lightweight) | Lightweight, no daemon. Cron handles scheduling. | Low |

---

## 1. Docker (Recommended)

The watcher ships with a `Dockerfile` and `docker-compose.yml`.

### One-command start

```bash
# 1. Copy and edit config
cp config.example.yaml config.yaml
# Fill in your Halo URL, credentials, Graph tenant/client info

# 2. Start
docker compose up -d

# 3. Verify
curl http://localhost:8888/health
# → {"status":"ok","conversations":0,"last_sync_at":"2026-08-12T..."}
```

### Custom port

```yaml
# docker-compose.yml
ports:
  - "9999:8888"   # Map host 9999 → container 8888
```

Then pass `--health-port 9999` to the container command.

### Logs

```bash
docker compose logs -f          # Follow logs
docker compose logs --tail=50   # Recent entries
```

### Updating

```bash
docker compose down
docker compose build --pull
docker compose up -d
```

---

## 2. Windows Service

Use **NSSM** (Non-Sucking Service Manager) to wrap the Python process
as a Windows service.

### Install NSSM

Download from [nssm.cc](https://nssm.cc/download) and extract. Add `nssm.exe`
to PATH or use the full path.

### Create the service

```powershell
# Install Python and dependencies
cd C:\halo-outlook-extension\watcher
pip install -e .

# Create a dedicated directory for the watcher
mkdir C:\halo-outlook-watcher
copy config.yaml C:\halo-outlook-watcher\

# Install the service
nssm install HaloWatcher python.exe
nssm set HaloWatcher AppDirectory "C:\halo-outlook-watcher"
nssm set HaloWatcher AppParameters "-m watcher.watcher --config config.yaml"
nssm set HaloWatcher DisplayName "Halo Outlook Watcher"
nssm set HaloWatcher Description "Journals Outlook conversations to HaloPSA tickets"
nssm set HaloWatcher Start SERVICE_AUTO_START

# Start
nssm start HaloWatcher
```

### Health check (Windows)

```powershell
Invoke-RestMethod http://localhost:8888/health
```

### Logs (Windows)

NSSM can redirect stdout/stderr to files:

```powershell
nssm set HaloWatcher AppStdout "C:\halo-outlook-watcher\stdout.log"
nssm set HaloWatcher AppStderr "C:\halo-outlook-watcher\stderr.log"
```

Or use the health endpoint to monitor liveness remotely.

---

## 3. Linux systemd

### Install

```bash
# Clone the repo
cd /opt
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension

# Install Python dependencies
cd watcher
pip install -e .
cd ..

# Copy and edit config
cp config.example.yaml /etc/halo-outlook-watcher/config.yaml
# Fill in your values
```

### systemd unit file

Create `/etc/systemd/system/halo-outlook-watcher.service`:

```ini
[Unit]
Description=Halo Outlook Extension — Watcher
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=halo-watcher
Group=halo-watcher
WorkingDirectory=/opt/halo-outlook-extension
ExecStart=/usr/bin/python3 -m watcher.watcher --config /etc/halo-outlook-watcher/config.yaml
Restart=always
RestartSec=10

# Resource limits
MemoryMax=256M
CPUQuota=50%

# Logging (journald)
StandardOutput=journal
StandardError=journal
SyslogIdentifier=halo-outlook-watcher

[Install]
WantedBy=multi-user.target
```

### Start

```bash
# Create the service user
sudo useradd --system --no-create-home halo-watcher

# Ensure permissions
sudo chown halo-watcher:halo-watcher /etc/halo-outlook-watcher/config.yaml
sudo chmod 600 /etc/halo-outlook-watcher/config.yaml

# Enable and start
sudo systemctl daemon-reload
sudo systemctl enable --now halo-outlook-watcher

# Verify
sudo systemctl status halo-outlook-watcher
curl http://localhost:8888/health
```

### Logs

```bash
sudo journalctl -u halo-outlook-watcher -f    # Follow
sudo journalctl -u halo-outlook-watcher -n 50 # Recent
```

---

## 4. Cron (Lightweight)

For users who don't want a persistent daemon, the watcher supports
`--once` mode. Cron handles scheduling — the watcher runs a single sync
cycle, then exits.

### Setup

```bash
# Install
cd /opt/halo-outlook-extension/watcher
pip install -e .

# Copy config
cp ../config.example.yaml /etc/halo-outlook-watcher/config.yaml
```

### Crontab

```cron
# Run every 2 minutes — the watcher syncs once and exits
*/2 * * * * /usr/bin/python3 -m watcher.watcher --once --config /etc/halo-outlook-watcher/config.yaml
```

### Tuning

The poll interval in `config.yaml` is irrelevant in `--once` mode (it only
affects the daemon). Instead, control frequency via the cron schedule:

| Schedule | Frequency | When to use |
|---|---|---|
| `*/1 * * * *` | Every minute | Near-real-time (more API calls) |
| `*/2 * * * *` | Every 2 minutes | Good balance |
| `*/5 * * * *` | Every 5 minutes | Lower API usage, acceptable delay |

### Overlap protection

The cron approach has no built-in overlap protection. If a sync cycle takes
longer than your interval, two instances may run simultaneously. Use
`flock` to prevent this:

```cron
*/2 * * * * /usr/bin/flock -n /tmp/halo-watcher.lock /usr/bin/python3 -m watcher.watcher --once --config /etc/halo-outlook-watcher/config.yaml
```

---

## Verifying Your Deployment

After deploying with any method, confirm it works:

```bash
# 1. Health check
curl http://localhost:8888/health
# Expected: {"status":"ok","conversations":0,"last_sync_at":"2026-08-12T..."}

# 2. Check logs for config validation
# Docker: docker compose logs
# systemd: sudo journalctl -u halo-outlook-watcher -n 20
# Cron: check your cron mail/log

# 3. Create a tracked conversation (via the add-in or manually)
# then verify messages appear as ticket actions in Halo
```

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| `halo.instance_url must start with https://` | Config URL has `http://` | Change to `https://` |
| `Halo connection failed` | Wrong credentials or URL | Run `setup_check.py` to validate |
| `Graph connection failed` | Wrong tenant_id, client_id, or client_secret | Verify app registration in Azure AD has `Mail.Read` permission, and admin consent is granted |
| Health endpoint unreachable | Firewall blocking port 8888 | Open the port or use `--health-port` with a different port |
| Docker health check failing | Container can't reach localhost health endpoint | Check `docker compose logs watcher` for startup errors |
| Messages not journaling | No conversations tracked yet | Use the Outlook add-in to create a ticket from a conversation, or manually insert a row into `state.db` |