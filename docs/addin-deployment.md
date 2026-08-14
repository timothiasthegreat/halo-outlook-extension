# Deploying the Add-in

The add-in is a set of static files served by the Express server in the Docker container. Build once, upload the manifest, done.

---

## Step 1: Build

```bash
cd add-in
npm install
npm run build
```

This produces `add-in/dist/` — HTML, JavaScript, icons. The Express server serves this directory. **No manual config editing needed** — the add-in fetches tenant config from `GET /api/config` at runtime.

---

## Step 2: Host

**With Docker (recommended):** the container already serves `add-in/dist/` via Express on port 3000. Nothing else needed.

**Without Docker:** run the Express server directly:

```bash
cd server
npm ci --production
node server.js --port 3000 --config ../config.yaml
```

---

## Step 3: Update the Manifest

Edit `add-in/manifest.xml`:

```xml
<SourceLocation DefaultValue="https://your-server:3000/taskpane.html"/>
```

Replace `your-server` with your actual hostname. The add-in's `/api/config` call uses `window.location.origin`, so matching the manifest URL to the Express server ensures it works.

---

## Step 4: Upload to Microsoft 365

1. Go to **Microsoft 365 Admin Center → Settings → Integrated Apps**
2. Click **Upload custom apps** → choose `manifest.xml`
3. Choose deployment scope (specific users or entire org)
4. Deploy

The add-in appears in Outlook within minutes.

---

## Step 5: First-Time User Setup

When a user first opens the add-in:

1. The add-in loads config from `/api/config`
2. If no PKCE token exists, it shows "Sign in to Halo"
3. After signing in, it shows the setup prompt: "Enable email watching?"
4. User clicks **Save & Start Watching** — their refresh token is encrypted and stored in `state.db`
5. The watcher begins syncing their conversations automatically

---

## Updating

Rebuild and restart:

```bash
cd add-in && npm run build && cd ..
docker compose restart
# Or: docker compose down && docker compose up -d
```

The add-in picks up the new build immediately — Outlook caches nothing. Config changes (action IDs, Halo URL) take effect on the next restart — no add-in rebuild needed.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Add-in shows "Not configured" | Express server must be running and reachable at the URL in manifest.xml |
| Config not loading | Check `GET /api/config` returns valid JSON |
| Setup button does nothing | Check `POST /api/register` is reachable from the add-in's origin |