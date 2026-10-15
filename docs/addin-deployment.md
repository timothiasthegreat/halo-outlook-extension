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
2. If no Halo token exists, it shows **Sign in to Halo**, which launches the PKCE auth flow
3. After sign-in, the add-in asks the server (`GET /api/register/status`) whether the watcher has a usable token for this mailbox. If not, it shows a yellow **Enable background sync** prompt
4. Clicking it runs a second Halo sign-in. That refresh token goes only to the watcher: the server checks the user's Halo access token and only accepts their own email, and the watcher confirms the refresh token belongs to that agent before encrypting it into `state.db`
5. The watcher syncs every conversation the user links from then on

The user's Halo agent email must match their Outlook mailbox address, or step 4 is refused with a 403. The prompt comes back whenever the watcher loses the token — Halo revoked it, or the state volume was wiped.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Add-in shows "Not configured" | Express server must be reachable at the URL in manifest.xml. Check `curl https://your-domain.com/health`. |
| Config not loading | Verify `GET /api/config` returns JSON. Check `docker compose logs halo-outlook`. |
| "Enable background sync" fails | Check `POST /api/register` is reachable from the add-in's origin. A 403 means the Halo agent's email doesn't match the Outlook mailbox address; fix the agent's email in Halo. |
| Add-in doesn't appear after upload | Wait up to 15 minutes. If still missing, check admin center deployment status. |
| Authentication fails | Verify `halo.client_id` in `config.yaml` is the Halo OAuth app's client ID (the add-in reads it from `/api/config`), the app uses the PKCE grant with scope `all`, and its redirect URI is `https://your-domain.com/auth-complete.html`. |
| "Background sync is off" keeps coming back | Halo rejected the watcher's token (HTTP 400/401 in the watcher logs as `token_refresh_failed`). Click **Enable background sync** again. Temporary Halo errors (429/5xx) don't cause this — they're retried. |
| Linked conversation shows as "not linked" | Tracking is per mailbox: the banner shows for the user who linked it, on any device. If it's missing for that user too, the watcher couldn't be reached or the conversation was unlinked — check `docker compose logs halo-outlook`. |