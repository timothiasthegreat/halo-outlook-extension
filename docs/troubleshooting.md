# Troubleshooting

Common problems and how to fix them.

---

## Watcher won't start

### `Config file not found: config.yaml`

The watcher can't find your configuration file.

```bash
# By default it looks for config.yaml in the working directory
cp config.example.yaml config.yaml
# Edit config.yaml with your values
python -m watcher.watcher --config config.yaml
```

### `halo.instance_url must start with https://`

Your Halo URL is missing `https://`. Change `http://` → `https://`.

### `Halo connection failed`

The watcher couldn't reach Halo or authenticate. Check:
- `halo.instance_url` is correct (no trailing slash, no /api path)
- `halo.client_id` is correct (`halo.client_secret` isn't used — PKCE needs no secret)
- Your Halo instance is reachable from where the watcher runs

### `Graph connection failed`

The watcher couldn't reach Microsoft Graph. Check:
- `graph.tenant_id`, `graph.client_id`, `graph.client_secret` are correct
- Admin consent has been granted for `Mail.Read` application permission
- The client secret hasn't expired

### `ModuleNotFoundError: No module named 'watcher'`

You haven't installed the watcher package.

```bash
cd watcher
pip install -e .
```

---

## Messages not journaling

### No conversations are being watched

The watcher's state database has no tracked conversations. This is normal if no emails have been linked yet.

**Check:** Use the add-in to create a ticket from a conversation, or manually insert a test row. `watched_by` must be a mailbox registered in `watched_mailboxes` with `token_status = 'active'` — rows without it are skipped:

```sql
-- In state.db (SQLite)
INSERT INTO conversations (conversation_id, ticket_id, created_at, last_sync_at, watched_by)
VALUES ('AAQkAD...', 1234, datetime('now'), datetime('now'), 'you@yourcompany.com');
```

### Background sync is off for a user

The watcher can only journal a mailbox it holds an active token for. Check:

```sql
SELECT email, token_status, last_token_refresh_at FROM watched_mailboxes;
```

- **`expired`** — Halo rejected the token (HTTP 400/401, logged as `token_refresh_failed`). The add-in shows **Enable background sync**; the user clicks it to register a fresh token.
- **No row** — the user never enabled background sync, or the state volume was wiped. Same fix.
- Temporary Halo problems (HTTP 429/5xx or network errors) don't expire tokens. The watcher logs them and retries on the next cycle.

Registration answers **403** when the user's Halo agent email doesn't match their Outlook mailbox address. Fix the agent's email in Halo.

### Watcher runs but doesn't find new messages

1. **Check the conversationId:** Verify the conversation exists in Exchange and has recent messages
2. **Check Graph permissions:** Run `setup_check.py` to confirm the token works and the mailbox is accessible
3. **Check dedup:** The watcher skips messages whose `internetMessageId` is already in the `synced_messages` table, or already on the ticket as an action's `email_message_id` (for example, the email the add-in journaled on Create Ticket). Messages synced before won't be re-synced
4. **Check staleness:** Conversations quiet for `stale_conversation_days` are marked stale. They're still watched, but only checked every `stale_poll_interval_seconds` (default hourly), so a late reply can take up to an hour to appear

### "The restriction or sort order is too complex" from Graph

Graph occasionally rejects the `$filter=conversationId eq '...'` query. The watcher retries automatically with exponential backoff. If it persists:

- Check the conversationId format — it should be a GUID-like string
- Try lowering `poll_interval_seconds` to reduce request frequency

### Action creation returns 401

API keys cannot create actions in Halo — the watcher posts with each user's own OAuth (PKCE) token. Verify `halo.client_id` belongs to an OAuth application (not an API key), and that the user's row in `watched_mailboxes` is `active` (see **Background sync is off for a user** above).

---

## Health check issues

### `curl http://localhost:8888/health` — connection refused

- The watcher must be running in daemon mode (not `--once`). The health server only starts in continuous poll mode.
- Check the watcher logs for `health_server_started`
- Verify the port isn't blocked by a firewall

### Docker health check failing

```bash
docker compose ps   # Check container status
docker compose logs watcher   # Check watcher logs
```

If the container starts but the health check fails, the watcher may have exited due to config validation. Check logs for `startup_aborted` or `config_validation_failed`.

---

## Add-in issues

### Add-in doesn't appear in Outlook

- The manifest must be uploaded to **Microsoft 365 Admin Center → Integrated Apps**
- It can take up to 24 hours to propagate (usually < 1 hour)
- The add-in only activates on message read (not compose). Open a received email.

### "Link to Halo" button does nothing

- The taskpane requires the add-in's static assets to be hosted at the URL in `manifest.xml`
- Check the `SourceLocation` in `manifest.xml` points to your hosted `taskpane.html`
- The hosting must be HTTPS (Office.js requirement)

### CORS errors in the add-in

The add-in sends Halo API calls through the Express server's proxy (`/api/proxy/*`), so those don't hit CORS. Only the sign-in token exchange (`/auth/token`) goes from the browser straight to Halo. If that fails with a CORS error, Halo must allow the add-in's origin.

### Conversation shows as "not linked"

The add-in asks the watcher whether a conversation is tracked each time you open it. Tracking belongs to the mailbox that linked it: the banner appears for that user on any device, but not for other users. If it's missing for the user who linked it, the watcher couldn't be reached (the add-in then falls back to what it cached on that device) or the conversation was unlinked.

### Create Ticket fails with 413

The email is bigger than a body-size limit. The Express server accepts up to 25 MB; a reverse proxy in front may allow less (nginx defaults to 1 MB). Set `client_max_body_size 25m;` in nginx.

### Can't find conversationId

Office.js provides `item.conversationId` on MessageRead items. If it's empty:
- The email must be opened in the reading pane
- Some Exchange configurations may not populate conversationId for very old emails

---

## State database issues

### `state.db` is missing or corrupted

If lost, restart the watcher — it auto-creates a fresh database. You do lose two things that live only in `state.db`:

- **Watcher tokens.** Each user clicks **Enable background sync** again (the add-in prompts them).
- **The watched-conversation list.** Conversations show as not linked until someone links them again.

Re-syncing is safe: before posting, the watcher checks the ticket's existing actions for each email's `internetMessageId`, so nothing is journaled twice.

### Database is locked

The watcher uses SQLite WAL mode for concurrent access. If you see "database is locked":
- Only one watcher process should write to the same database
- If using cron mode, ensure overlap protection (`flock`) is configured

---

## Performance

### High CPU usage

- Lower `poll_interval_seconds` if you have many watched conversations
- The watcher makes one Graph API call per conversation per cycle
- With 100+ watched conversations, consider increasing the interval to 180-300s

### High memory usage

The watcher is designed to use < 100 MB. If memory grows:
- Check for memory leaks in the Python process (unlikely with async/await)
- The `synced_messages` table can grow large — run `prune_stale_synced()` periodically

---

## Getting help

1. **Check the logs:** The watcher logs structured output to stdout. Set `log_level: DEBUG` for verbose logging.
2. **Run setup_check.py:** Validates config, Halo connectivity, Graph connectivity, and action IDs.
3. **Check the health endpoint:** `curl http://localhost:8888/health` shows conversations watched and last sync time.
4. **Open an issue:** [GitHub Issues](https://github.com/timothiasthegreat/halo-outlook-extension/issues)