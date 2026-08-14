# Configuration Reference

Everything in `config.yaml`, explained.

> **File location:** `config.yaml` in the repository root (or pass `--config <path>` to the watcher).
> **Template:** Copy `config.example.yaml` → `config.yaml` and fill in your values.

---

## Top-level structure

```yaml
halo:      # HaloPSA connection
graph:     # Microsoft Graph connection
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

The watcher appends `/auth` (OAuth token endpoint) and `/api` (REST API base) automatically. Do not include trailing paths.

### `halo.client_id`

**Required.** OAuth2 client ID from HaloPSA.

Where to find it: **Configuration → Integrations → Halo API → OAuth Application**. Create an application with `client_credentials` grant type and the `all` scope.

```yaml
halo:
  client_id: "your-halo-client-id"
```

### `halo.client_secret`

**Required.** OAuth2 client secret from the same Halo OAuth application.

```yaml
halo:
  client_secret: "your-halo-client-secret"
```

### `halo.actions`

**Required.** Ticket Action outcome IDs for your Halo instance. These control what happens when an action is posted:

- `email_received` — used when a customer emails you (inbound). Default: `0` ("Email Update").
- `email_sent` — used when you email the customer (outbound). Default: `16` ("Email User").
- `internal_note` — used for internal journal entries. Default: `7` ("Recorded Note").

```yaml
halo:
  actions:
    email_received: 0
    email_sent: 16
    internal_note: 7
```

**These IDs are instance-specific.** Every Halo deployment may have different IDs. Run the setup check script to auto-discover yours:

```bash
python scripts/setup_check.py --discover-actions
```

Or find them manually in **Configuration → Tickets → Ticket Actions**.

### `halo.custom_field_conv_id`

**Required.** Numeric ID of the custom field that stores the Exchange `conversationId` on tickets.

```yaml
halo:
  custom_field_conv_id: 285
```

This field is how the system knows which tickets are tracked to which conversations. See [Halo Setup](halo-setup.md) for instructions on creating this field.

### `halo.default_ticket_type_id`

**Optional.** Default ticket type for new tickets created from conversations. Default: `1` (Incident).

```yaml
halo:
  default_ticket_type_id: 1
```

Adjust to match your Halo instance's ticket type IDs: Incident, Service Request, Change, Problem, Sales, etc.

---

## `graph` — Microsoft Graph Connection

### `graph.tenant_id`

**Required.** Your Azure AD tenant ID (GUID).

```yaml
graph:
  tenant_id: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

Find it in **Azure Portal → Azure Active Directory → Overview → Tenant ID**.

### `graph.client_id`

**Required.** Application (client) ID from your Azure AD app registration.

```yaml
graph:
  client_id: "d4c3b2a1-f6e5-0987-dcba-0987654321ef"
```

See [Azure Setup](azure-setup.md) for step-by-step app registration instructions.

### `graph.client_secret`

**Required.** Client secret from the same app registration.

```yaml
graph:
  client_secret: "your-graph-client-secret"
```

Create this in **Azure Portal → App Registrations → Your App → Certificates & Secrets → New Client Secret**.

### `graph.user_email`

**Required.** The mailbox to watch. This is the user's personal email address — the one customers email directly.

```yaml
graph:
  user_email: you@yourcompany.com
```

The watcher polls `GET /users/{user_email}/messages` to find new messages in watched conversations. This must match the mailbox the user reads in Outlook.

---

## `watcher` — Watcher Behavior

All keys under `watcher` are **optional** with sensible defaults.

### `watcher.poll_interval_seconds`

How often the watcher checks for new messages. Default: `90` seconds. Minimum: `30`, maximum: `3600`.

```yaml
watcher:
  poll_interval_seconds: 90
```

- **Lower** (30-60s) → faster sync, more API calls, lower latency
- **Higher** (180-300s) → fewer API calls, higher latency, lower resource usage
- In `--once` (cron) mode this is ignored — the cron schedule controls frequency

### `watcher.stale_conversation_days`

Stop watching conversations that haven't had any activity in N days. Default: `14`.

```yaml
watcher:
  stale_conversation_days: 14
```

Minimum: 1 day, maximum: 365 days. Stale conversations are marked inactive but not deleted — they can be reactivated if a new message arrives.

### `watcher.log_level`

Log verbosity. Default: `INFO`.

```yaml
watcher:
  log_level: INFO
```

Options: `DEBUG`, `INFO`, `WARNING`, `ERROR`. Use `DEBUG` when troubleshooting — it logs every message inspected and action created.

### `watcher.state_db_path`

Path to the SQLite state database. Default: `state.db` in the working directory.

```yaml
watcher:
  state_db_path: /var/lib/halo-watcher/state.db
```

This file tracks which messages have been synced and which conversations are being watched. It's not critical — if lost, the watcher re-syncs recent messages (dedup logic prevents duplicates).

---

## Health Check Port

The watcher's health endpoint listens on `localhost:8888` by default. Change it with `--health-port`:

```bash
python -m watcher.watcher --health-port 9999
```

In Docker, map the port in `docker-compose.yml`:

```yaml
ports:
  - "9999:8888"
```

Then pass `--health-port 8888` (the container port, not the host port).

---

## Complete Example

```yaml
halo:
  instance_url: https://your-instance.halopsa.com
  client_id: "halo-oauth-client-id"
  client_secret: "halo-oauth-client-secret"
  actions:
    email_received: 0
    email_sent: 16
    internal_note: 7
  custom_field_conv_id: 285
  default_ticket_type_id: 1

graph:
  tenant_id: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
  client_id: "d4c3b2a1-f6e5-0987-dcba-0987654321ef"
  client_secret: "graph-client-secret"
  user_email: you@yourcompany.com

watcher:
  poll_interval_seconds: 90
  stale_conversation_days: 14
  log_level: INFO
  state_db_path: state.db
```