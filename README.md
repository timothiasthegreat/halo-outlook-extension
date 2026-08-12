# Halo Outlook Extension

> Create and track HaloPSA tickets directly from Outlook. Click a button to link any email conversation to a ticket — every future reply automatically journals, even when Outlook isn't running.

**One click in Outlook → ticket exists in Halo.** Customer replies journal as ticket updates. Your replies journal as agent actions. The customer only ever sees your personal email address — no shared mailbox jarring.

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

# Host the dist/ folder (GitHub Pages, Azure Static Web Apps, or any web server)
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
| [Deployment](docs/deployment.md) | Docker, Windows Service, Linux systemd, cron |
| [Troubleshooting](docs/troubleshooting.md) | Common errors, log inspection, health checks |

---

## Requirements

### Watcher
- Python 3.11+
- HaloPSA (self-hosted or Halo-hosted) with API access
- Azure AD app registration with `Mail.Read` application permission
- Always-on environment (Docker, Windows Service, Linux systemd, or cron)

### Add-in
- Microsoft 365 tenant with Outlook
- Node.js 18+ for building
- Static web hosting (GitHub Pages, Azure, any web server)

### HaloPSA
- A **custom field** (text, ticket-scoped) to store `conversationId`
- OAuth2 application credentials (client ID + secret)
- Ticket Action IDs for "Email Received" and "Email Sent" (discoverable via setup script)

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
4. **Journals to Halo** → Posts ticket action with correct `outcome_id`
5. **OnMessageSend** → Add-in intercepts outbound sends, journals instantly (future enhancement; currently watcher handles it)

---

## Community

This project is MIT-licensed and intended for the HaloPSA community. Deploy it on your own infrastructure — no dependency on the original author's systems.

- **Issues & PRs:** [GitHub](https://github.com/timothiasthegreat/halo-outlook-extension)
- **Discussion:** Open an issue for questions, feature requests, or deployment help
- **Contributing:** PRs welcome — please run `python -m pytest tests/` before submitting

---

## License

MIT © [Fireside IT Partners](https://firesideit.ca)