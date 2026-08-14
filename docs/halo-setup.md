# HaloPSA Setup

Configuring your Halo instance for the Outlook Extension.

---

## 1. Create the Conversation ID Custom Field

1. Go to **Configuration → Custom Fields** → **New**
2. Fill in:
   - **Name:** `Ticket Conversation ID`
   - **System Name:** `ticketconvid` (alphanumeric only — no underscores/hyphens)
   - **Type:** Text (single line)
   - **Scope:** Tickets
3. Save
4. Add the field to each ticket type you want to track: **Configuration → Tickets → Ticket Types → [type] → Custom Fields tab**

Find the field's numeric ID:

```bash
python scripts/setup_check.py --list-custom-fields
```

Set in `config.yaml`:

```yaml
halo:
  custom_field_conv_id: 285   # your field's ID
```

---

## 2. Discover Ticket Action IDs

These are **instance-specific**. Auto-discover:

```bash
python scripts/setup_check.py --discover-actions
```

Typical output:

```
email_received → 0 (Email Update)
email_sent     → 16 (Email User)
internal_note  → 7 (Recorded Note)
```

---

## 3. Create the OAuth2 Application (PKCE)

You need **one OAuth app** for both the add-in and watcher.

1. Go to **Configuration → Integrations → Halo API → OAuth Applications**
2. Click **New**
3. Fill in:
   - **Name:** `Outlook Extension`
   - **Grant Type:** Auth Code (PKCE) — NOT Client Credentials
   - **Scope:** `all`
   - **Redirect URI:** `https://your-server:3000/auth-complete.html`
4. Save. Copy the **Client ID** only — no client secret is needed for PKCE.

Set in `config.yaml`:

```yaml
halo:
  client_id: "the-client-id"
  # client_secret: "" — not needed
```

Set in the add-in's manifest: update `manifest.xml` with the same OAuth app client ID.

---

## Why PKCE instead of client credentials?

- **Client credentials** always attribute to a single "Automation" agent. Multi-user orgs get confusing ticket timelines.
- **PKCE** gives each user their own refresh token. The watcher uses that user's token when journaling their messages — actions are correctly attributed to the human.
- **Fewer secrets to manage.** No client secret in config. No second OAuth app. One app, one client ID.

---

## Verification

```bash
python scripts/setup_check.py
```

Expected:

```
✓ config.yaml loaded and valid
✓ HaloPSA reachable
✓ Action IDs discovered
```

---

## Reference: Halo API Behavior

| Behavior | Detail |
|---|---|
| POST body must be JSON array | `[{...}]` not `{...}` |
| Ticket update uses POST | `POST /api/Tickets` with `[{"id": N, ...}]` |
| Custom field system names | Alphanumeric only. No underscores. `CFticketconvid` ✓ |
| Custom field filtering | NOT supported via API — watcher uses local state.db |
| Actions require OAuth | API keys can't create actions (401) — OAuth required |
| OAuth apps are agent-bound | PKCE per-user tokens give correct attribution per agent |