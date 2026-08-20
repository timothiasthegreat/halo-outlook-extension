# Configuration Reference

Everything in `config.yaml`, explained.

> **File location:** `config.yaml` in the repository root.
> **Template:** Copy `config.example.yaml` → `config.yaml`.

---

## Top-level structure

```yaml
halo:      # HaloPSA connection (PKCE — no client secret needed)
graph:     # Microsoft Graph connection (client credentials)
watcher:   # Watcher behavior (optional — has defaults)
```

---

## `halo` — HaloPSA Connection

### `halo.instance_url`

**Required.** Your HaloPSA instance URL. Must start with `https://`.

```yaml
halo:
  instance_url: https://your-instance.halopsa.com
```

### `halo.client_id`

**Required.** OAuth2 client ID from HaloPSA. This is the **same OAuth app** used by the add-in.

Where to find it: **Configuration → Integrations → Halo API → OAuth Application**. Create an application with **Auth Code (PKCE)** grant type and the `all` scope. No client secret is needed — the add-in handles the PKCE flow.

```yaml
halo:
  client_id: "your-halo-client-id"
```

### `halo.client_secret`

**Optional.** Default: `""` (empty). Not needed for PKCE auth.

```yaml
halo:
  client_secret: ""   # leave empty — PKCE doesn't use it
```

### `halo.actions`

Ticket Action outcome IDs. These are instance-specific — run the discovery script to find yours:

```bash
python scripts/setup_check.py --discover-actions
```

```yaml
halo:
  actions:
    email_received: 0    # inbound: customer → agent
    email_sent: 16       # outbound: agent → customer
    internal_note: 7     # internal journal
```

### `halo.custom_field_conv_id`

Numeric ID of the custom field storing `conversationId`. Default: `285`.

```yaml
halo:
  custom_field_conv_id: 285
```

### `halo.default_ticket_type_id`

Default ticket type for new tickets. Default: `1` (Incident).

---

### `halo.exclusions`

Control which ticket types and statuses are visible in the add-in UI.
All fields default to empty lists (no exclusions). The entire block is
optional — omit it entirely for default behaviour.

Exclusions only affect the add-in UI — they do **not** stop the watcher
from syncing already-linked conversations. Once a ticket is linked to a
conversation, sync continues regardless of later exclusion changes.

```yaml
halo:
  exclusions:                       # optional — entire block can be omitted
    ticket_type_ids_create: []      # optional — defaults to []
    ticket_type_ids_search: []      # optional — defaults to []
    status_ids_search: []           # optional — defaults to []
```

| Field | Type | Default | Description |
|---|---|---|---|
| `ticket_type_ids_create` | `list[int]` | `[]` | Ticket type IDs to exclude from the Create Ticket dropdown. Useful for types that shouldn't be created from Outlook (e.g., Quotes, Change Requests). |
| `ticket_type_ids_search` | `list[int]` | `[]` | Ticket type IDs to exclude from search results. |
| `status_ids_search` | `list[int]` | `[]` | Ticket status IDs to exclude from search results. Useful for hiding closed/resolved tickets from the link dialog. |

---

## `graph` — Microsoft Graph Connection

### `graph.tenant_id`

**Required.** Azure AD tenant ID (GUID).

### `graph.client_id`

**Required.** App registration client ID. Requires `Mail.Read` application permission with admin consent.

### `graph.client_secret`

**Required.** Client secret from the same app registration.

### `graph.user_email`

**Optional.** Default: `""` (empty). Per-user mailboxes are registered via the add-in's PKCE flow — leave this blank for multi-user deployments. Set it for single-user deployments where you want a fixed default.

---

## `watcher` — Watcher Behavior

All optional with sensible defaults.

| Key | Default | Range | Description |
|---|---|---|---|
| `poll_interval_seconds` | `90` | 30–3600 | How often to check for new messages |
| `stale_conversation_days` | `14` | 1–365 | Auto-unwatch after N days of silence |
| `log_level` | `INFO` | DEBUG/INFO/WARNING/ERROR | Log verbosity |
| `state_db_path` | `state.db` | any path | SQLite database location |

---

## Environment Variables

Set these outside `config.yaml`:

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `FERNET_KEY` | Production | Auto-generated in Docker | 44-char base64 key for encrypting refresh tokens at rest |
| `FERNET_KEY_FILE` | No | `/app/state/fernet.key` | Path to persisted key file |

---

## Ports

| Port | Service | Purpose |
|---|---|---|
| `3000` | Express | Static add-in + `/api/config` + `/api/register` |
| `8888` | Watcher | Health check (`GET /health`) |

---

## Complete Example

```yaml
halo:
  instance_url: https://your-instance.halopsa.com
  client_id: "halo-oauth-client-id"
  # client_secret: "" — not needed for PKCE
  actions:
    email_received: 0
    email_sent: 16
    internal_note: 7
  custom_field_conv_id: 285
  default_ticket_type_id: 1

graph:
  tenant_id: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
  client_id: "graph-app-client-id"
  client_secret: "graph-client-secret"
  # user_email: "" — registered per-user via add-in

watcher:
  poll_interval_seconds: 90
  stale_conversation_days: 14
  log_level: INFO
  state_db_path: state.db
```