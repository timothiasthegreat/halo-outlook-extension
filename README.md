# Halo Outlook Extension

> Create and track HaloPSA tickets directly from Outlook. Click a button to link any email conversation to a ticket — every future reply automatically journals, even when Outlook isn't running.

**One container, one deploy.** The add-in is served from the same Docker container that runs the watcher. Register your mailbox by clicking one button in Outlook — no config editing, no second OAuth app, no build-time values to fill in.

---

## Prerequisites

| What | Required for | Must have |
|---|---|---|
| **HaloPSA instance** | Both | Admin access to Configuration → Integrations → Halo API. HTTPS required. |
| **OAuth2 app** (PKCE auth code flow) | Both | A **single** OAuth Application in Halo. Client ID only — no client secret needed. Use the same app for both the add-in and watcher. |
| **Custom field** on tickets | Both | Text-type ticket-scoped custom field to store conversationId. System name alphanumeric only (e.g. `ticketconvid`). Add to every ticket type you track. |
| **Ticket Action IDs** | Both | Numeric IDs for "Email Received" and "Email Sent". Run `scripts/setup_check.py --discover-actions` to find yours. |
| **Azure AD app registration** | Watcher | Application permission `Mail.Read` with admin consent. |
| **Exchange Online mailboxes** | Watcher | Licensed Exchange Online mailboxes. Users self-register via the add-in — no per-user config. |
| **Always-on environment** | Both | Docker (recommended), or bare-metal Python 3.11 + Node.js 18. |

### Before you start

```bash
cp config.example.yaml config.yaml   # fill in halo.client_id + graph.*
python scripts/setup_check.py        # validates connectivity
```

---

## How It Works

**One container** serves everything:

| Component | Tech | What it does |
|---|---|---|
| **Express Server** | Node.js (in container) | Serves the add-in's static files + `GET /api/config` + `POST /api/register` |
| **Outlook Add-in** | TypeScript + React (Office.js) | "Create Ticket" / "Link to Ticket" buttons. Fetches config at runtime — no rebuild for changes. |
| **Watcher Service** | Python 3.11 (in container) | Polls Graph for new messages in watched conversations, journals to Halo. Uses per-user PKCE tokens for correct action attribution. |

**Multi-user by design:** Each user authorizes once via the add-in's PKCE flow. Their refresh token is stored encrypted in `state.db`. The watcher uses that user's token when journaling messages from their mailbox — actions are correctly attributed to the human who owns the conversation.

---

## Quick Start

### 1. Clone and configure

```bash
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension

# Copy and edit the config (halo.client_id + graph.* are the only required fields)
cp config.example.yaml config.yaml
```

### 2. Build and run

```bash
# Build the add-in
cd add-in && npm install && npm run build && cd ..

# Start the combined container
docker compose up -d
```

Verify: `curl http://localhost:3000/health` and `curl http://localhost:3000/api/config`

### 3. Deploy the add-in

```bash
# The add-in is already served from the container — just upload the manifest.
# Update manifest.xml <SourceLocation> to your hosted URL (https://your-server:3000)
# Upload manifest.xml to Microsoft 365 Admin Center → Integrated Apps
```

The add-in appears in Outlook within minutes. The first time a user opens it, they'll be prompted to register their mailbox — one click, done.

### 4. Start tracking

1. Open any customer email in Outlook
2. Click **"Link to Halo"** in the reading pane
3. Choose **Create Ticket** or **Link to Ticket**
4. All future replies automatically journal — zero manual steps

---

## Environment Variables

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `FERNET_KEY` | Production only | Auto-generated in Docker | Encrypts refresh tokens at rest. Docker entrypoint generates one on first run and persists it to `/app/state/fernet.key`. |
| `FERNET_KEY_FILE` | No | `/app/state/fernet.key` | Path to persisted Fernet key. |
| `EXPRESS_PORT` | No | `3000` | Express listen port. |
| `HEALTH_PORT` | No | `8888` | Watcher health-check port. |
| `PYTHONUNBUFFERED` | No | — | Set to `1` for unbuffered logs. |

---

## Documentation

| Doc | What's covered |
|---|---|
| [Configuration](docs/configuration.md) | Every `config.yaml` option |
| [Halo Setup](docs/halo-setup.md) | Creating the custom field, discovering action IDs, PKCE OAuth app |
| [Azure Setup](docs/azure-setup.md) | App registration, permissions, admin consent |
| [Add-in Deployment](docs/addin-deployment.md) | Building, hosting, uploading the manifest |
| [Watcher Deployment](docs/deployment.md) | Docker, bare-metal, cron |
| [Troubleshooting](docs/troubleshooting.md) | Common errors, logs, health checks |

---

## Architecture

```
┌──────────┐     ┌──────────────────────┐     ┌──────────┐
│  Outlook  │────▶│   Express Server     │────▶│ HaloPSA  │
│ (user UI) │     │ (static + API)       │     │ REST API │
└──────────┘     │                      │     └──────────┘
                 │ /api/config          │           ▲
┌──────────┐     │ /api/register        │           │
│ Exchange  │────▶│ /health              │───────────┘
│ (Graph)   │     │                      │
└──────────┘     │ Watcher Service      │
                 │ (Python daemon)      │
                 │ TokenManager         │
                 │ (per-user PKCE)      │
                 └───────┬──────────────┘
                         │
                  ┌──────▼──────┐
                  │  state.db   │
                  │  (SQLite)   │
                  │  WAL mode   │
                  └─────────────┘
```

1. **User opens add-in** → Fetches config from `/api/config`, presents "Create Ticket" or "Link to Ticket"
2. **First-time setup** → User clicks "Enable Watching", PKCE refresh token is encrypted and stored in `state.db` via `/api/register`
3. **Watcher polls Graph** → Loads watched mailboxes from `state.db`, fetches new messages per-user mailbox
4. **Dedup & direction** → Skips already-synced messages, determines inbound vs outbound per user's email
5. **Journals to Halo** → Posts action using that user's PKCE token — correctly attributed to the human
6. **Staleness detection** → Idle conversations auto-marked stale after `watcher.stale_conversation_days`

---

## Community

MIT-licensed. Deploy on your own infrastructure — no dependency on the original author's systems.

- **Issues & PRs:** [GitHub](https://github.com/timothiasthegreat/halo-outlook-extension)
- **Discussion:** Open an issue for questions, feature requests, or deployment help
- **Contributing:** PRs welcome — run `python -m pytest watcher/tests/` and `cd server && node --test tests/` before submitting

---

## License

MIT © [Fireside IT Partners](https://firesideit.ca)