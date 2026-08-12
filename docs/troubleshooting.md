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
- `halo.client_id` and `halo.client_secret` are correct
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

**Check:** Use the add-in to create a ticket from a conversation, or manually insert a test row:

```sql
-- In state.db (SQLite)
INSERT INTO conversations (conversation_id, ticket_id, created_at, last_sync_at)
VALUES ('AAQkAD...', 1234, datetime('now'), datetime('now'));
```

### Watcher runs but doesn't find new messages

1. **Check the conversationId:** Verify the conversation exists in Exchange and has recent messages
2. **Check Graph permissions:** Run `setup_check.py` to confirm the token works and the mailbox is accessible
3. **Check dedup:** The watcher skips messages with `internetMessageId` already in `synced_messages` table. If messages were already synced before, they won't be re-synced

### "The restriction or sort order is too complex" from Graph

Graph occasionally rejects the `$filter=conversationId eq '...'` query. The watcher retries automatically with exponential backoff. If it persists:

- Check the conversationId format — it should be a GUID-like string
- Try lowering `poll_interval_seconds` to reduce request frequency

### Action creation returns 401

API keys cannot create actions in Halo. You must use OAuth2 client credentials. Verify your `halo.client_id` and `halo.client_secret` correspond to an OAuth application (not an API key).

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

The add-in calls Halo's API directly from the browser. If you see CORS errors:
- Halo must allow the add-in's origin. Some Halo instances may need CORS configuration.
- Alternative: proxy API calls through the add-in's hosting server

### Can't find conversationId

Office.js provides `item.conversationId` on MessageRead items. If it's empty:
- The email must be opened in the reading pane
- Some Exchange configurations may not populate conversationId for very old emails

---

## State database issues

### `state.db` is missing or corrupted

The state database is not critical. If lost, simply restart the watcher — it will auto-create a fresh database. Previously synced messages may be re-synced, but the dedup logic prevents duplicate journal entries.

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