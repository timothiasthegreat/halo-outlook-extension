# HaloPSA Setup

Configuring your Halo instance for the Outlook Extension.

---

## 1. Create the Conversation ID Custom Field

This field stores the Exchange `conversationId` on each ticket. It's what the watcher uses to find tracked conversations.

### In Halo:

1. Go to **Configuration → Custom Fields**
2. Click **New**
3. Fill in:
   - **Name:** `Ticket Conversation ID` (or whatever you prefer)
   - **System Name:** `ticketconvid` (alphanumeric only — no underscores, no hyphens)
   - **Type:** Text (single line)
   - **Scope:** Tickets
   - **Usage:** Both (visible to agents and end-users, or agent-only)
4. Save

> **Important:** The system name must be alphanumeric. Halo silently rejects custom field names with underscores or hyphens. `ticketconvid` ✓ — `ticket_conv_id` ✗.

### Add the field to your ticket types:

For each ticket type you want to track:

1. Go to **Configuration → Tickets → Ticket Types**
2. Select a ticket type (e.g., Incident, Sales)
3. Click the **Custom Fields** tab
4. Add your new `ticketconvid` field

> Without this step, `GET /api/Tickets/{id}` won't return the custom field value — even if it's set on the ticket.

### Find the field's numeric ID:

1. Go back to **Configuration → Custom Fields**
2. Find your `ticketconvid` field in the list
3. The numeric ID is shown in the URL when you edit the field, or in the API response from `GET /api/customfields`

```bash
# Or use the setup check script:
python scripts/setup_check.py --list-custom-fields
```

Set this ID as `halo.custom_field_conv_id` in `config.yaml`:

```yaml
halo:
  custom_field_conv_id: 285  # Replace with your field's ID
```

---

## 2. Discover Ticket Action IDs

Halo uses **Ticket Actions** (also called outcome IDs) to control what happens when you post an action to a ticket. Each action has a numeric ID — and these are **different on every Halo instance**.

The two actions this project needs:

| Concept | Common name | Default IDs (may differ) |
|---|---|---|
| Email received from customer | "Email Update" | 0 |
| Email sent to customer | "Email User" | 16 |
| Internal note | "Recorded Note" | 7 |

### Auto-discovery (recommended):

```bash
python scripts/setup_check.py --discover-actions
```

This queries your Halo instance's action list and matches on name. Output:

```
Discovered actions:
  email_received → 0 (Email Update)
  email_sent     → 16 (Email User)
  internal_note  → 7 (Recorded Note)
```

### Manual discovery:

1. Go to **Configuration → Tickets → Ticket Actions**
2. Find "Email Update" (or equivalent) — note the ID
3. Find "Email User" (or equivalent) — note the ID
4. Set in `config.yaml`:

```yaml
halo:
  actions:
    email_received: <email-update-id>
    email_sent: <email-user-id>
```

---

## 3. Create OAuth2 Application

The watcher needs unattended API access. Halo supports OAuth2 client credentials for this.

1. Go to **Configuration → Integrations → Halo API**
2. Click **OAuth Applications**
3. Click **New**
4. Fill in:
   - **Name:** `Outlook Watcher`
   - **Grant Type:** Client Credentials
   - **Scope:** `all` (required — the watcher reads tickets, creates actions, and updates custom fields)
   - **Client ID:** Copy the generated value
   - **Client Secret:** Copy the generated value (you won't see it again)
5. Save

Set these in `config.yaml`:

```yaml
halo:
  client_id: "the-client-id"
  client_secret: "the-client-secret"
```

> **Note:** API keys (`X-Halo-Api-Key` header) also work for reading tickets and creating them, but they **cannot create ticket actions** (POST /api/Actions returns 401). OAuth2 client credentials are required for the watcher to post journal entries. The add-in uses auth code flow with PKCE for user-attributed actions.

---

## Verification

Run the setup check to confirm everything is wired correctly:

```bash
python scripts/setup_check.py
```

Expected output:

```
✓ config.yaml loaded and valid
✓ HaloPSA reachable → https://your-instance.halopsa.com
✓ OAuth2 token acquired (scope: all)
✓ Custom field 285 (CFticketconvid) found
✓ Action: email_received → 0 (Email Update)
✓ Action: email_sent → 16 (Email User)
All checks passed — watcher is ready to start.
```

---

## Reference: Halo API Behavior

Known quirks that matter for this project:

| Behavior | Detail |
|---|---|
| POST body must be JSON array | `[{...}]` not `{...}` |
| Ticket update uses POST | `POST /api/Tickets` with `[{ "id": N, ... }]` — PATCH/PUT return 405 |
| Custom field system names | Alphanumeric only. No underscores. `CFticketconvid` ✓, `CFticket_conv_id` ✗ |
| Custom field filtering | NOT supported via API. The watcher uses a local SQLite state store instead. |
| Actions require OAuth | API keys cannot create actions (401). OAuth2 client credentials required. |