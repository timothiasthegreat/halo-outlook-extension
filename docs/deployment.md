# Deploying the Halo Outlook Extension

The add-in server and watcher ship as a single prebuilt Docker image:
**`docker.io/timothiasthegreat/halo-outlook-extension:latest`**.

You provide two files tailored to your environment — `config.yaml` and `manifest.xml` — and deploy. The image never needs a local build.

---

## What you need

| Artifact | You provide | Details |
|---|---|---|
| `config.yaml` | ✅ One per environment | HaloPSA and Graph credentials. Template: `config.example.yaml`. |
| `manifest.xml` | ✅ One per environment | Add-in manifest with your domain's URLs. Template: `add-in/manifest.xml`. |
| Docker image | ❌ Prebuilt | `docker.io/timothiasthegreat/halo-outlook-extension:latest` |
| Reverse proxy with TLS | ✅ One per environment | Nginx, Caddy, Traefik — Outlook requires HTTPS. |

---

## Step 1: Create `config.yaml`

Copy the template and fill in your values:

```bash
cp config.example.yaml config.yaml
```

**Required fields:**

```yaml
halo:
  instance_url: https://your-instance.halopsa.com
  client_id: "your-pkce-client-id"   # OAuth app from HaloPSA (PKCE, scope=all)

graph:
  tenant_id: "a1b2c3d4-..."         # Azure AD tenant ID
  client_id: "graph-app-client-id"   # App registration with Mail.Read
  client_secret: "graph-secret"      # Client secret
```

Everything else has sensible defaults. See [Configuration Reference](configuration.md) for all options.

---

## Step 2: Customize `manifest.xml`

The manifest in the repo (`add-in/manifest.xml`) contains `@source_location@` placeholders. **You must replace these** with your domain before deploying the add-in to Microsoft 365.

Find-and-replace every `@source_location@` with your server's URL:

```
https://your-domain.com
```

The result should look like:

```xml
<IconUrl DefaultValue="https://your-domain.com/assets/icon-64.png" />
<SourceLocation DefaultValue="https://your-domain.com/taskpane.html" />
```

**After customizing, validate:**

```bash
npx office-addin-manifest validate manifest.xml
```

See the [Add-in Deployment Guide](addin-deployment.md) for the full upload process.

---

## Step 3: Deploy the container

The image has two volume mount points:

| Container path | Purpose | Back up? |
|---|---|---|
| `/app/config.yaml` | Your configuration file (read-only) | No — it's in version control |
| `/app/state` | Persistent `state.db` + `fernet.key` | **Yes** — contains encrypted user tokens |

### Docker Compose (recommended)

```yaml
version: "3.8"

services:
  halo-outlook:
    image: docker.io/timothiasthegreat/halo-outlook-extension:latest
    container_name: halo-outlook-extension
    restart: unless-stopped
    expose:
      - "3000"   # Internal — the reverse proxy routes to this
    volumes:
      - ./config.yaml:/app/config.yaml:ro
      - watcher_state:/app/state
    environment:
      - PYTHONUNBUFFERED=1
      - FERNET_KEY_FILE=/app/state/fernet.key

  # TLS-terminating reverse proxy (required — Outlook enforces HTTPS)
  caddy:
    image: caddy:2-alpine
    restart: unless-stopped
    ports:
      - "443:443"
      - "80:80"
    volumes:
      - ./Caddyfile:/etc/caddy/Caddyfile:ro
      - caddy_data:/data

volumes:
  watcher_state:
  caddy_data:
```

**Caddyfile:**

```
your-domain.com {
    reverse_proxy halo-outlook:3000
}
```

### Portainer stack

Same compose file. In Portainer:

1. **Stacks → Add stack** → paste the compose above
2. Set `config.yaml` and `Caddyfile` paths to absolute host paths (e.g. `/opt/halo-outlook/config.yaml`)
3. Deploy

### Nginx (alternative)

If you already have an nginx reverse proxy:

```nginx
server {
    listen 443 ssl;
    server_name your-domain.com;

    ssl_certificate     /etc/ssl/certs/your-domain.crt;
    ssl_certificate_key /etc/ssl/private/your-domain.key;

    location / {
        proxy_pass http://halo-outlook:3000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

## Step 4: Verify

```bash
curl https://your-domain.com/health        # {"status":"ok"}
curl https://your-domain.com/api/config    # returns your config
curl https://your-domain.com/taskpane.html # returns HTML
```

All three should return 200. If `/api/config` fails, your `config.yaml` isn't mounted or has a syntax error.

---

## Step 5: Upload the add-in to Microsoft 365

1. Go to **Microsoft 365 Admin Center → Settings → Integrated Apps**
2. **Upload custom apps** → select your customized `manifest.xml`
3. Assign to users
4. Wait 5–15 minutes for propagation

For local testing, you can sideload instead: [Add-in Deployment → Sideloading](addin-deployment.md).

---

## Updating

Pull the latest image and restart:

```bash
docker compose pull halo-outlook
docker compose up -d
```

Or in Portainer: **Re-pull image** → **Update the stack**.

The `watcher_state` volume persists across updates — user registrations survive image changes.

---

## Backups

Only one directory needs backing up: the `watcher_state` volume. It contains:

- `state.db` — all mailbox registrations, watched conversations, sync state
- `fernet.key` — encryption key for refresh tokens at rest

If this volume is lost, every user must re-register their mailbox through the add-in.

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| `config.yaml` not found | Volume mount path wrong | Verify absolute path on host: `ls /opt/halo-outlook/config.yaml` |
| `halo.instance_url must start with https://` | Config has `http://` | Change to `https://` |
| `Graph connection failed` | Wrong tenant ID or secret | Verify app registration has `Mail.Read` + admin consent |
| Add-in shows "Not configured" | `/api/config` not reachable | Check Express is running: `docker compose logs halo-outlook` |
| Add-in fails to load in Outlook | Missing or misconfigured TLS | Verify reverse proxy has valid cert; check `npx office-addin-manifest validate` |
| Messages not journaling | No users registered | User must open add-in and click "Enable Watching" |
| `FERNET_KEY environment variable is required` | Key not generated | Docker entrypoint auto-generates on first run. Check `/app/state/fernet.key` exists in the volume. |
| Portainer stack fails to pull | Docker Hub rate limiting | Log in to Docker Hub in Portainer: **Registries → Add registry** → `docker.io` with your credentials |