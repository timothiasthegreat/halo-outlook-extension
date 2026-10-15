# Halo Outlook Extension

> Create and track HaloPSA tickets directly from Outlook. Click a button to link any email conversation to a ticket — every future reply automatically journals, even when Outlook isn't running.

**One container, one deploy.** The add-in is served from the same Docker container that runs the watcher. Each user turns on background sync with one click in Outlook — no per-user config, no second OAuth app, no build-time values to fill in.

---

## Prerequisites

| What | Required for | Must have |
|---|---|---|
| **HaloPSA instance** | Both | Admin access to Configuration → Integrations → Halo API. HTTPS required. |
| **OAuth2 app** (PKCE auth code flow) | Both | A **single** OAuth Application in Halo. Client ID only — no client secret needed. Use the same app for both the add-in and watcher. It must issue refresh tokens. |
| **Custom field** on tickets | Both | Text-type ticket-scoped custom field to store conversationId. System name alphanumeric only (e.g. `ticketconvid`). Add to every ticket type you track. |
| **Ticket Action IDs** | Both | Numeric IDs for "Email Received" and "Email Sent". Run `scripts/setup_check.py --discover-actions` to find yours. |
| **Azure AD app registration** | Watcher | Application permission `Mail.Read` with admin consent. |
| **Exchange Online mailboxes** | Watcher | Licensed Exchange Online mailboxes. Users self-register via the add-in — no per-user config. |
| **Matching agent emails** | Both | Each user's Halo agent email must be the same address as their Outlook mailbox. Registration is refused (403) otherwise. |
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
| **Express Server** | Node.js (in container) | Serves the add-in's static files, `GET /api/config`, the Halo API proxy, and the mailbox/conversation endpoints (`/api/register`, `/api/conversations`). Every endpoint that touches watcher state requires the user's Halo access token. |
| **Outlook Add-in** | TypeScript + React (Office.js) | "Create Ticket" / "Link to Ticket" buttons. Fetches config at runtime — no rebuild for changes. |
| **Watcher Service** | Python 3.11 (in container) | Polls Graph for new messages in watched conversations, journals to Halo. Uses per-user PKCE tokens for correct action attribution. |

**Multi-user by design:** Each user signs in to Halo in the add-in, then clicks **Enable background sync**, which runs a second, one-time Halo sign-in. That sign-in's refresh token goes only to the watcher and is stored encrypted in `state.db`; the add-in keeps its own. Because the two never share a token, Halo rotating one can't break the other. The watcher uses each user's token when journaling messages from their mailbox — actions are correctly attributed to the human who owns the conversation.

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

The add-in appears in Outlook within minutes. The first time a user opens it, they click **Sign in to Halo**, then **Enable background sync** (a second, one-time Halo sign-in that gives the watcher its own token). The prompt comes back by itself if the watcher ever loses that token.

### 4. Start tracking

1. Open any customer email in Outlook
2. Click **"Link to Halo"** in the reading pane
3. Choose **Create Ticket** or **Link to Ticket**
4. All future replies automatically journal — zero manual steps

The banner shows the linked ticket on any device where you open the conversation. **Unlink** stops journaling and clears the conversation ID from the ticket in Halo.

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
│ Exchange  │────▶│ /api/conversations   │───────────┘
│ (Graph)   │     │ /api/proxy/*         │
└──────────┘     │ /health              │
                 │                      │
                 │ Watcher Service      │
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

1. **User opens add-in** → Fetches config from `/api/config`, asks the watcher (`GET /api/conversations`) whether this conversation is tracked, and shows the ticket banner (refreshed from Halo) or "Create Ticket" / "Link to Ticket"
2. **First-time setup** → User clicks "Enable background sync". This runs a separate Halo sign-in whose refresh token goes only to the watcher (the add-in keeps its own). `/api/register` checks the caller's Halo token, and the watcher confirms the refresh token belongs to that agent before encrypting it into `state.db`
3. **Watcher polls Graph** → Loads watched conversations from `state.db` and fetches new messages (and their attachments) from each watching user's mailbox, following every page of results
4. **Dedup & direction** → Skips messages already recorded in `state.db` or already journaled on the ticket in Halo — so the email the add-in journals on Create Ticket, or a re-sync after `state.db` is lost, never posts twice — and determines inbound vs outbound per user's email
5. **Journals to Halo** → Posts action using that user's PKCE token — correctly attributed to the human
6. **Staleness detection** → Idle conversations marked stale after `watcher.stale_conversation_days`; they stay linked, are checked hourly (`stale_poll_interval_seconds`) and reactivate on the next reply

---

## Security

- **Authenticated endpoints.** `/api/register`, `/api/register/status` and `/api/conversations` require the caller's Halo access token. The server asks Halo who the token belongs to, and callers can only register and watch their own mailbox.
- **Verified tokens.** Before storing a refresh token, the watcher exchanges it and checks with Halo that it belongs to the agent being registered.
- **Encrypted at rest.** Refresh tokens are Fernet-encrypted in `state.db`, with the key in `fernet.key` (or `FERNET_KEY`).
- **Outages don't log users out.** A token is only marked expired when Halo rejects it (HTTP 400/401). Rate limits, server errors and network failures are retried on the next cycle.
- **Internal API stays internal.** The watcher's API on port 8888 is only for the Express server inside the container. Don't publish that port.

---

## Community

MIT-licensed. Deploy on your own infrastructure — no dependency on the original author's systems.

- **Issues & PRs:** [GitHub](https://github.com/timothiasthegreat/halo-outlook-extension)
- **Discussion:** Open an issue for questions, feature requests, or deployment help
- **Contributing:** PRs welcome — run `python -m pytest watcher/tests/` and `cd server && npm test` before submitting. The server tests read `config.yaml` from the repo root, so copy `config.example.yaml` to `config.yaml` first.

---

## License

MIT © [Fireside IT Partners](https://firesideit.ca)