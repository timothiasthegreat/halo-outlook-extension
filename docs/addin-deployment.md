# Add-in Deployment

The add-in is prebuilt into the Docker image. You only need to customize `manifest.xml` with your domain and upload it.

---

## Step 1: Customize the manifest

The source manifest at `add-in/manifest.xml` uses `@source_location@` placeholders throughout. **Replace every occurrence** with your server's HTTPS URL:

```
https://your-domain.com
```

All URLs in the manifest must use HTTPS — Outlook enforces this.

### What gets replaced

| Placeholder | Becomes |
|---|---|
| `@source_location@/taskpane.html` | `https://your-domain.com/taskpane.html` |
| `@source_location@/assets/icon-64.png` | `https://your-domain.com/assets/icon-64.png` |
| `@source_location@/assets/icon-128.png` | `https://your-domain.com/assets/icon-128.png` |
| `@source_location@/assets/icon-16.png` | `https://your-domain.com/assets/icon-16.png` |
| (and 25, 32, 48, 80 sizes) | |

### Validate after customizing

```bash
npx office-addin-manifest validate manifest.xml
```

Acceptance test must show "Package Type Identified" and "Correct Package." The XML schema warning about `VersionOverrides` namespace is a known validator divergence — it does not block store submission.

---

## Step 2: Upload to Microsoft 365

1. Go to **Microsoft 365 Admin Center → Settings → Integrated Apps**
2. Click **Upload custom apps** → choose your customized `manifest.xml`
3. Select deployment scope:
   - **Specific users/groups** — recommended for initial rollout
   - **Entire organization** — if everyone needs it
4. Click **Deploy**

The add-in appears in Outlook within **5–15 minutes**. Users will see "Halo Ticket Bridge" in the ribbon when reading an email.

---

## Sideloading (local testing)

For testing before tenant-wide deployment, sideload directly in Outlook Desktop:

1. Open Outlook Desktop
2. Ribbon → **Get Add-ins** → **My Add-ins**
3. **Add a custom add-in** → **Add from File**
4. Select your customized `manifest.xml`

The manifest URLs must point to a reachable server. For local development, use `webpack serve` (which handles HTTPS):

```bash
cd add-in
npm start   # starts on https://localhost:3000
```

Then use a manifest with `https://localhost:3000` URLs.

> **Note:** `npm start` runs webpack-dev-server locally. The Docker image does not need this — the Express server inside the container serves the prebuilt add-in at port 3000.

---

## Per-user setup

When a user first opens the add-in:

1. The add-in loads tenant config from `GET /api/config` (no rebuild needed)
2. If no Halo token exists, it shows "Sign in to Halo" — launches the PKCE auth flow
3. After authenticating, the user clicks **Save & Start Watching**
4. Their refresh token is encrypted and stored in `state.db`
5. The watcher begins syncing their conversations automatically

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Add-in shows "Not configured" | Express server must be reachable at the URL in manifest.xml. Check `curl https://your-domain.com/health`. |
| Config not loading | Verify `GET /api/config` returns JSON. Check `docker compose logs halo-outlook`. |
| Setup button does nothing | Check `POST /api/register` is reachable from the add-in's origin. |
| Add-in doesn't appear after upload | Wait up to 15 minutes. If still missing, check admin center deployment status. |
| Authentication fails | Verify the Halo OAuth app client ID in `config.yaml` matches the one in the manifest's HaloPSA setup. Same client ID, same scope (`all`), same grant type (PKCE). |