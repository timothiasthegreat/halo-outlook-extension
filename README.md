# Halo Outlook Extension

> Create and track HaloPSA tickets directly from Outlook. Click a button to link any email conversation to a ticket — every future reply automatically journals, even when Outlook isn't running.

**One click in Outlook → ticket exists in Halo.** Customer replies journal as ticket updates. Your replies journal as agent actions. The customer only ever sees your personal email address — no shared mailbox jarring.

---

## Prerequisites

You need all of the following before you can deploy:

| What | Required for | Must have |
|---|---|---|
| **HaloPSA instance** with REST API | Both components | Admin access to Configuration → Integrations → Halo API. Must be reachable over HTTPS. Self-hosted or Halo‑hosted both work. |
| **OAuth2 client credentials** in Halo | Watcher service | Client ID + client secret from an OAuth Application (grant type: Client Credentials, scope: `all`). API keys **cannot** post ticket actions — OAuth is required. |
| **Custom field** on Halo tickets | Both components | A text-type custom field (ticket‑scoped) to store Exchange `conversationId`. Must be added to every ticket type you want to track. System name must be alphanumeric only (e.g. `ticketconvid`). |
| **Ticket Action IDs** | Both components | Numeric IDs for "Email Received" (inbound) and "Email Sent" (outbound). Usually `0` and `16` — verify with `scripts/setup_check.py --discover-actions`. |
| **Azure AD app registration** | Watcher service | Application permission `Mail.Read` granted with admin consent. Provides the `graph.tenant_id`, `graph.client_id`, and `graph.client_secret` for `config.yaml`. |
| **Exchange Online mailbox** | Watcher service | A licensed Exchange Online mailbox whose email address goes in `graph.user_email`. This is the address the user reads in Outlook — the one customers actually send to. |
| **Python 3.11+** | Watcher service | On the machine that runs the watcher (Docker, Windows, or Linux). |
| **Always‑on environment** | Watcher service | Docker, Windows Service, Linux systemd, or cron. The watcher must run continuously to capture messages when Outlook is closed. |
| **Microsoft 365 tenant** | Outlook add-in | Your org's tenant. The add-in is uploaded to Admin Center → Integrated Apps. |
| **Node.js 18+** | Outlook add-in | For building the add-in's static assets (one‑time: `npm install && npm run build`). Not needed at runtime. |
| **Static web hosting** | Outlook add-in | Any HTTPS web server to host the built `dist/` folder. GitHub Pages, Azure Static Web Apps, S3, or self‑hosted all work. |

### Before you start — run the pre‑flight check

```bash
cp config.example.yaml config.yaml   # fill in your values
python scripts/setup_check.py        # validates everything
```

If it passes, you're ready. If not, the output tells you exactly what's missing.

---

## How It Works

Two components, one repository:

| Component | Tech | What it does |
|---|---|---|
| **Outlook Add-in** | TypeScript + React (Office.js) | "Create Ticket" and "Link to Ticket" buttons in the Outlook reading pane. Shows ticket status banner on tracked emails. |
| **Watcher Service** | Python 3.11+ (async) | Background daemon that polls Microsoft Graph for new messages in tracked conversations and journals them to Halo. Runs even when Outlook is closed. |

The magic: Exchange `conversationId` stored in a Halo custom field. When a conversation is linked, the watcher sees every new message in that thread and pushes it to the correct ticket with the correct action type (inbound vs outbound).

---

## Quick Start

### 1. Clone and configure

```bash
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension

# Copy and edit the config
cp config.example.yaml config.yaml
# Fill in your Halo URL, credentials, Graph tenant info, and action IDs
```

### 2. Run the watcher

```bash
# Docker (recommended — one command)
docker compose up -d

# Or bare-metal
cd watcher && pip install -e .
python -m watcher.watcher --config ../config.yaml
```

Verify it's running: `curl http://localhost:8888/health`

### 3. Deploy the add-in

```bash
cd add-in

# Edit src/config.ts with your Halo URL and action IDs
# Build
npm install && npm run build

# Host the dist/ folder (GitHub Pages, Azure, S3, or any web server)
# Full guide: docs/addin-deployment.md
# Update manifest.xml <SourceLocation> to your hosted URL
# Upload manifest.xml to Microsoft 365 Admin Center → Integrated Apps
```

The add-in appears in Outlook within minutes.

### 4. Start tracking

1. Open any customer email in Outlook
2. Click **"Link to Halo"** in the reading pane
3. Choose **Create Ticket** or **Link to Ticket**
4. All future replies automatically journal — zero manual steps

---

## Documentation

| Doc | What's covered |
|---|---|
| [Configuration](docs/configuration.md) | Every `config.yaml` option explained, with defaults |
| [Halo Setup](docs/halo-setup.md) | Creating the custom field, discovering action IDs |
| [Azure Setup](docs/azure-setup.md) | App registration, permissions, admin consent |
| [Add-in Deployment](docs/addin-deployment.md) | Building, hosting, and uploading the Outlook add-in |
| [Watcher Deployment](docs/deployment.md) | Docker, Windows Service, Linux systemd, cron |
| [Troubleshooting](docs/troubleshooting.md) | Common errors, log inspection, health checks |

---

## Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│   Outlook   │────▶│  Outlook Add-in  │────▶│  HaloPSA    │
│  (user UI)  │     │  (Office.js)     │     │  REST API   │
└─────────────┘     └──────────────────┘     └─────────────┘
                                                      ▲
┌─────────────┐     ┌──────────────────┐              │
│  Exchange   │────▶│  Watcher Service │──────────────┘
│  (Graph API)│     │  (Python daemon) │
└─────────────┘     └──────────────────┘
                           │
                     ┌─────▼──────┐
                     │  state.db  │
                     │  (SQLite)  │
                     └────────────┘
```

1. **User clicks "Create Ticket" in Outlook** → Add-in calls Halo API, creates ticket with `conversationId` in custom field
2. **Watcher polls Graph API** → Finds new messages in watched conversations
3. **Dedup & direction** → Skips already-synced messages, determines inbound vs outbound
4. **Journals to Halo** → Posts ticket action with correct `outcome_id`; downloads and attaches file attachments automatically
5. **Staleness detection** → Conversations idle for longer than `watcher.stale_conversation_days` are automatically un-watched

---

## Community

This project is MIT-licensed and intended for the HaloPSA community. Deploy it on your own infrastructure — no dependency on the original author's systems.

- **Issues & PRs:** [GitHub](https://github.com/timothiasthegreat/halo-outlook-extension)
- **Discussion:** Open an issue for questions, feature requests, or deployment help
- **Contributing:** PRs welcome — please run `python -m pytest tests/` before submitting

---

## License

MIT © [Fireside IT Partners](https://firesideit.ca)