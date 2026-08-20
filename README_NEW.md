# Halo Outlook Extension

[![Docker Pulls](https://img.shields.io/docker/pulls/timothiasthegreat/halo-outlook-extension)](https://hub.docker.com/r/timothiasthegreat/halo-outlook-extension)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/timothiasthegreat/halo-outlook-extension/actions/workflows/docker-push.yml/badge.svg)](https://github.com/timothiasthegreat/halo-outlook-extension/actions)

> **Turn Outlook email conversations into HaloPSA tickets — automatically.**
> Click once to link a conversation. Every future reply journals to the ticket,
> even when Outlook isn't running.

---

## What Problem Does This Solve?

HaloPSA users who handle sales, account management, or relationship-driven work
live in Outlook. Customers email them directly (`you@yourcompany.com` rather than
`help@yourcompany.com`). These conversations should become Halo tickets so
they're trackable, searchable, and auditable — the same as support tickets.

**The friction every Halo user faces:**

- Switching to the Halo web UI to create and update tickets is a workflow tax
  nobody pays consistently.
- Halo's built-in mailbox monitoring creates tickets from _every_ email —
  newsletters, vendor spam, internal threads. You don't want all of those.
- Using a shared mailbox as the sender for ticket replies means the customer sees
  a different address than the person they emailed. That breaks relationship
  continuity.

**What this extension gives you:**

1. Click **"Create Ticket"** on any email in Outlook — a ticket is created, and
   the initial email is journaled.
2. Reply from Outlook normally. Every subsequent message — your sends and the
   customer's replies — journals to the ticket **automatically**.
3. The customer only ever sees your personal email. No shared mailbox needed.
4. If you archive the thread or Outlook isn't running, messages still get
   captured by the background watcher.

---

## Screenshots

<!-- TODO: add screenshot of task pane showing "Create Ticket" / "Link to Ticket" buttons -->
<!-- TODO: add screenshot of tracked-email banner showing ticket reference, status, Open in Halo link -->

---

## How It Works

```
┌──────────┐     ┌──────────────────────┐     ┌──────────┐
│  Outlook  │────▶│   Express Server     │────▶│ HaloPSA  │
│ (user UI) │     │ (static + API)       │     │ REST API │
└──────────┘     │                      │     └──────────┘
                 │ /api/config          │           ▲
┌──────────┐     │ /api/register        │           │
│ Exchange  │────▶│ /api/conversations  │───────────┘
│ (Graph)   │     │                      │
└──────────┘     │ Watcher Service      │
                 │ (Python daemon)      │
                 │ Per-user PKCE tokens │
                 └───────┬──────────────┘
                         │
                  ┌──────▼──────┐
                  │  state.db   │
                  │  (SQLite)   │
                  │  WAL mode   │
                  └─────────────┘
```

1. **User opens the add-in** in the reading pane → "Create Ticket" or "Link to
   Ticket" based on whether the conversation is already tracked.
2. **First-time setup** → user authenticates via a popup OAuth dialog (PKCE).
   Their refresh token is encrypted and stored in `state.db`.
3. **Watcher polls Microsoft Graph** every 90 seconds → fetches new messages for
   every watched conversation, per-user mailbox.
4. **Dedup & direction** → skips already-synced messages, determines inbound vs.
   outbound, and posts ticket actions attributed to the correct user.

### Multi-User by Design

Every user authenticates independently via the add-in's PKCE flow. Their refresh
token is encrypted at rest (Fernet) and used by the watcher when journaling
messages from **their** mailbox. Ticket actions are correctly attributed to the
human who owns the conversation — not a service account.

**No shared mailbox. No client secret for Halo.** Each user keeps their own
email identity, and the watcher uses their individual token.

### PKCE OAuth Flow

The add-in uses a pure browser-based PKCE authorization code flow:

1. User clicks "Sign in" → Office dialog opens to `halo_url/auth/authorize`
2. User logs in and authorizes the app
3. Halo redirects to `auth-complete.html` with the authorization code
4. The add-in exchanges the code for an access + refresh token at
   `halo_url/auth/token`
5. The refresh token is stored in localStorage (for add-in API calls) **and**
   forwarded to the watcher's `/api/register` endpoint (for background
   journaling)
6. The watcher encrypts the refresh token with Fernet and stores it in `state.db`

The watcher later decrypts the token, exchanges it for access tokens via
`refresh_token` grant, and handles token rotation automatically.

---

## Before You Start

### You Will Need

| What | Required For | Where to Get It |
|---|---|---|
| **HaloPSA instance** (HTTPS) | Add-in + watcher API calls | Your Halo instance URL |
| **Halo OAuth2 Application** (PKCE, `all` scope) | PKCE auth code flow | Configuration → Integrations → Halo API |
| **Halo custom field** (text, ticket-scoped) | Storing conversationId on tickets | Configuration → Custom Fields |
| **Ticket Action IDs** (email_received, email_sent) | Correct journal action types | Run `scripts/setup_check.py --discover-actions` |
| **Azure AD App Registration** + `Mail.Read` application permission | Watcher polls Graph for messages | Azure Portal → App Registrations |
| **Exchange Online licensed mailboxes** | Graph needs real mailboxes to query | Each user needs a license |
| **Docker** (or Python 3.11 + Node.js 18 bare metal) | Runs the combined container | [Docker Desktop](https://www.docker.com/products/docker-desktop/) |

### Pre-Flight Checklist

```bash
# Clone and configure
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension
cp config.example.yaml config.yaml

# Fill in these values in config.yaml:
#   halo.instance_url, halo.client_id
#   graph.tenant_id, graph.client_id, graph.client_secret

# Validate connectivity
pip install httpx pyyaml pydantic
python scripts/setup_check.py                     # full pre-flight
python scripts/setup_check.py --discover-actions  # find your action IDs
```

Once the pre-flight passes, you're ready to install.

---

## Installation

### Option 1: Prebuilt Docker Image (Recommended)

```bash
# Pull the image (CI builds on every push to main)
docker pull timothiasthegreat/halo-outlook-extension:latest

# Start the container
docker run -d \
  --name halo-outlook \
  -p 3000:3000 \
  -v $(pwd)/config.yaml:/app/config.yaml:ro \
  -v halo-state:/app/state \
  timothiasthegreat/halo-outlook-extension:latest
```

The container starts Express (port 3000, serving the add-in + API) and the
Python watcher (health check on internal port 8888). The Fernet encryption key
is auto-generated on first run and persisted to the `halo-state` volume.

### Option 2: Build from Source

```bash
git clone https://github.com/timothiasthegreat/halo-outlook-extension.git
cd halo-outlook-extension

# Configure
cp config.example.yaml config.yaml
# Edit config.yaml with your values

# Build the add-in static assets
cd add-in && npm install && npm run build && cd ..

# Start
docker compose up -d
```

The `docker-compose.yml` mounts `./config.yaml` read-only and a persistent
`watcher_state` volume at `/app/state` (holds `state.db` + `fernet.key`).

### Verify

```bash
curl http://localhost:3000/health        # → {"status":"ok"}
curl http://localhost:3000/api/config    # → your tenant config object
docker logs halo-outlook-extension       # watcher + Express logs
```

---

## Deploy the Add-in to Outlook

The add-in's static HTML/JS/CSS is already served from the Express container at
`http://your-server:3000`. You just need to tell Outlook where to find it.

### 1. Update the manifest

Open `add-in/dist/manifest.xml` (generated by the build step). Replace every
instance of `@source_location@` with your hosted URL:

```xml
<!-- Before -->
<SourceLocation DefaultValue="@source_location@/taskpane.html" />

<!-- After -->
<SourceLocation DefaultValue="https://haloext.yourcompany.com/taskpane.html" />
```

Also update the `<AppDomains>` section — replace the placeholder Halo URL with
your actual Halo instance URL:

```xml
<AppDomain>https://your-instance.halopsa.com</AppDomain>
```

### 2. Upload to Microsoft 365

1. Go to **Microsoft 365 Admin Center** → **Settings** → **Integrated Apps**
2. Click **Upload custom apps** → upload `manifest.xml`
3. Wait ~5 minutes. The **"Link to Halo"** button appears in Outlook's reading
   pane ribbon.

### Hosting options

| Approach | When to Use |
|---|---|
| **Same container** (default) | Single-server setup. Express at port 3000 serves everything. |
| **Reverse proxy** | Put nginx/Caddy in front of port 3000 with TLS — e.g. `https://haloext.yourcompany.com` → `localhost:3000`. |
| **Static hosting** | GitHub Pages, Azure Static Web Apps, S3 + CloudFront. Requires separate watcher deployment. |

---

## Daily Use

### Creating a Ticket

1. Open any customer email in Outlook
2. Click **"Link to Halo"** in the reading pane ribbon
3. Select a ticket type from the dropdown (loaded dynamically from your Halo
   instance)
4. Click **"Create Ticket"**
5. The task pane updates to show the ticket reference, status, assigned agent,
   and a clickable **"Open in Halo ↗"** link

The add-in automatically registers your mailbox with the watcher and tracks the
conversation — no separate setup step needed.

### Linking to an Existing Ticket

1. Click **"Link to Halo"**
2. Click **"Link to Ticket"** button
3. Type a ticket reference number or customer name and press Enter
4. Select the target ticket from the search results
5. All future replies in this conversation journal to that ticket

### What Happens Automatically

- **You reply from Outlook** → the watcher picks up your reply via Graph and
  posts it as an email-sent action on the ticket.
- **The customer replies** → posted as an email-received action.
- **Attachments on replies** → downloaded and attached to the ticket (up to 10
  MB per file).
- **Conversation goes quiet for 14 days** → marked stale (configurable via
  `watcher.stale_conversation_days`). The mapping is kept but syncing pauses.

### Stopping Tracking

Click the **"Unlink"** button in the task pane to remove the conversation→ticket
mapping from `state.db`. This doesn't delete the ticket — it just stops
automatic journaling for that conversation.

---

## Configuration Reference

Full reference: [docs/configuration.md](docs/configuration.md)

| Key | Required | Default | Description |
|---|---|---|---|
| `halo.instance_url` | **Yes** | — | Your HaloPSA URL (must start with `https://`) |
| `halo.client_id` | **Yes** | — | PKCE OAuth client ID (same app used by add-in) |
| `halo.client_secret` | No | `""` | Not needed for PKCE — leave empty |
| `halo.actions.email_received` | **Yes** | `0` | Action ID for inbound email |
| `halo.actions.email_sent` | **Yes** | `16` | Action ID for outbound email |
| `halo.actions.internal_note` | No | `7` | Action ID for internal notes |
| `halo.custom_field_conv_id` | **Yes** | `285` | Numeric ID of conversationId custom field |
| `halo.default_ticket_type_id` | No | `1` | Default ticket type for new tickets |
| `graph.tenant_id` | **Yes** | — | Azure AD tenant GUID |
| `graph.client_id` | **Yes** | — | App registration client ID |
| `graph.client_secret` | **Yes** | — | App registration client secret |
| `graph.user_email` | No | `""` | Default mailbox (leave blank — users self-register) |
| `watcher.poll_interval_seconds` | No | `90` | How often to check for new messages |
| `watcher.stale_conversation_days` | No | `14` | Auto-mark stale after N days of silence |
| `watcher.log_level` | No | `INFO` | `DEBUG`, `INFO`, `WARNING`, or `ERROR` |
| `watcher.state_db_path` | No | `state.db` | SQLite database location |

---

## Environment Variables

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `FERNET_KEY` | Production | Auto-generated | 44-char base64 key for encrypting refresh tokens at rest |
| `FERNET_KEY_FILE` | No | `/app/state/fernet.key` | Path to persisted key file |
| `EXPRESS_PORT` | No | `3000` | Express listen port |
| `HEALTH_PORT` | No | `8888` | Watcher health-check port (internal only) |
| `PYTHONUNBUFFERED` | No | — | Set to `1` for unbuffered log output |

The Docker entrypoint auto-generates `FERNET_KEY` on first run and persists it
to the volume. You only need to set it explicitly for bare-metal deployments.

---

## Troubleshooting

### Health checks

```bash
# Is the container running?
docker ps | grep halo-outlook

# Express (port 3000)
curl http://localhost:3000/health
# → {"status":"ok"}

# Watcher logs (includes sync cycle summaries)
docker logs halo-outlook-extension --tail 50

# Follow live logs
docker logs -f halo-outlook-extension
```

### Common issues

| Symptom | Likely Cause | Fix |
|---|---|---|
| **"Link to Halo" button doesn't appear in Outlook** | Manifest not deployed, or `@source_location@` wasn't replaced | Re-upload manifest with correct URL, wait 5 min |
| **Add-in shows "Not authenticated" loop** | Halo OAuth app misconfigured, wrong client ID, or redirect URI mismatch | Verify client ID in `halo.client_id` matches the OAuth app |
| **Watcher logs show "token_refresh_failed"** | User's refresh token expired or was revoked | User re-authenticates in the add-in (re-registers mailbox) |
| **"Graph connection failed" in startup logs** | Graph credentials wrong or consent missing | Run `scripts/setup_check.py`, verify admin consent granted |
| **Messages not journaling** | Conversation not in `state.db`, or watcher not running | `curl localhost:8888/health` to verify watcher is alive; check `state.db` for conversation records |

### Manual state.db inspection

```bash
# Inside the container
docker exec -it halo-outlook-extension sqlite3 /app/state/state.db

# Check registered mailboxes
SELECT email, token_status, last_token_refresh_at FROM watched_mailboxes;

# Check tracked conversations
SELECT conversation_id, ticket_id, watched_by, last_sync_at FROM conversations WHERE is_stale = 0;
```

---

## Architecture

```
┌──────────┐     ┌──────────────────────┐     ┌──────────┐
│  Outlook  │────▶│   Express Server     │────▶│ HaloPSA  │
│ (user UI) │     │ (static + API)       │     │ REST API │
└──────────┘     │                      │     └──────────┘
                 │ /api/config          │           ▲
┌──────────┐     │ /api/register        │           │
│ Exchange  │────▶│ /api/conversations  │───────────┘
│ (Graph)   │     │ /api/proxy/*        │
└──────────┘     │ /api/ticket-types    │
                 │ /health              │          ┌──────────────┐
                 │                      │─────────▶│ Docker Hub   │
                 │ Watcher Service      │          │ (CI builds)  │
                 │ (Python 3.11 async)  │          └──────────────┘
                 │ TokenManager         │
                 │ (per-user PKCE)      │
                 └───────┬──────────────┘
                         │
                  ┌──────▼──────┐
                  │  state.db   │
                  │  (SQLite)   │
                  │  ├─ watched_mailboxes
                  │  ├─ conversations
                  │  └─ synced_messages
                  └─────────────┘
```

**Components:**

| Component | Language | Role |
|---|---|---|
| **Express Server** | Node.js 18 | Serves add-in static files + API endpoints. Proxies state.db writes to the watcher's FastAPI server (port 8888) to avoid multi-process SQLite corruption. |
| **Outlook Add-in** | TypeScript + React | Task pane UI: create/link tickets, auth dialog, ticket banner. Uses Office.js APIs for conversation context. |
| **Watcher Service** | Python 3.11 | Polls Graph API for new messages. Creates per-user `HaloClient` instances for correct action attribution. Journals attachments. Manages PKCE token lifecycle. |
| **state.db** | SQLite (WAL) | Tracks conversation→ticket mappings, synced message dedup, and encrypted per-user refresh tokens. |
| **Docker Entrypoint** | Bash | Generates Fernet key on first run, starts Express in background, runs watcher in foreground. |

---

## Documentation

| Doc | Content |
|---|---|
| [Configuration](docs/configuration.md) | Every config.yaml option explained |
| [Halo Setup](docs/halo-setup.md) | Creating the custom field, discovering action IDs, PKCE OAuth app |
| [Azure Setup](docs/azure-setup.md) | App registration, permissions, admin consent |
| [Add-in Deployment](docs/addin-deployment.md) | Manifest, hosting options, M365 upload |
| [Deployment](docs/deployment.md) | Watcher deployment: Docker, bare-metal, cron |
| [Troubleshooting](docs/troubleshooting.md) | Full error catalog with solutions |

---

## Development

```bash
# Python watcher tests
cd watcher && pip install -e ".[dev]" && pytest -v

# Node.js server tests
cd server && npm install && npm test

# TypeScript add-in build
cd add-in && npm install && npm run build
```

CI builds and pushes the Docker image to
[`timothiasthegreat/halo-outlook-extension`](https://hub.docker.com/r/timothiasthegreat/halo-outlook-extension)
on every push to `main`.

---

## Community

MIT-licensed. Deploy on your own infrastructure — no dependency on the original
author's systems.

- **Issues & PRs:** [GitHub](https://github.com/timothiasthegreat/halo-outlook-extension)
- **Docker Hub:** [`timothiasthegreat/halo-outlook-extension`](https://hub.docker.com/r/timothiasthegreat/halo-outlook-extension)
- **Questions:** Open a GitHub issue — feature requests and deployment help welcome.

---

## License

MIT © [Fireside IT Partners](https://firesideit.ca)