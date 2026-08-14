# Deploying the Add-in

The Outlook add-in is a set of static files (HTML, JavaScript, CSS, icons) that Outlook loads into the reading pane. You build once, host anywhere, and never touch again until you want to update.

---

## Step 1: Configure

Edit `add-in/src/config.ts` with your Halo instance's values:

```ts
const config: AddinConfig = {
  haloUrl: "https://your-instance.halopsa.com",   // Your Halo URL
  haloClientId: "",                             // OAuth2 client ID
  actions: {
    emailReceived: 0,                           // Verify with setup_check.py
    emailSent: 16,
  },
  customFieldConvId: 285,                       // Your custom field ID
  defaultTicketTypeId: 1,                       // Default ticket type
};
```

All other files are generic — this is the only file you need to edit.

---

## Step 2: Build

```bash
cd add-in
npm install
npm run build
```

Output: `add-in/dist/` containing:

```
dist/
├── taskpane.html        # The pane Outlook loads
├── taskpane.js          # Bundled React app
├── manifest.xml         # Copied from root (SourceLocation still needs updating)
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── ...              # Replace these with your own logos
```

Replace the placeholder icon PNGs in `add-in/assets/` with your own before building (or swap them in `dist/assets/` after). Icons must be square, minimum 16×16 through 128×128.

---

## Step 3: Update the manifest

Open `dist/manifest.xml` (or `add-in/manifest.xml`). Replace every occurrence of `@source_location@` with the URL where your `dist/` folder will be hosted:

```xml
<!-- Before -->
<SourceLocation DefaultValue="@source_location@/taskpane.html" />
<bt:Image id="icon16" DefaultValue="@source_location@/assets/icon-16.png" />

<!-- After -->
<SourceLocation DefaultValue="https://your-domain.com/halo-addin/taskpane.html" />
<bt:Image id="icon16" DefaultValue="https://your-domain.com/halo-addin/assets/icon-16.png" />
```

Search for `@source_location@` — it appears in ~14 places. Replace all of them with your URL. The URL must be **HTTPS** (Office.js requirement).

---

## Step 4: Host

Upload the entire `dist/` folder to any static web host that supports HTTPS.

### GitHub Pages (free)

```bash
# Create a dedicated repo or use a gh-pages branch
cd add-in/dist
git init
git add -A
git commit -m "Add-in static assets"
git branch -M gh-pages
git remote add origin https://github.com/your-org/halo-outlook-addin.git
git push -u origin gh-pages
```

Your add-in will be at `https://your-org.github.io/halo-outlook-addin/taskpane.html`.

Enable GitHub Pages in the repo settings: **Settings → Pages → Source: Deploy from a branch → gh-pages → / (root)**.

### Azure Static Web Apps

```bash
# Install SWA CLI
npm install -g @azure/static-web-apps-cli

# Deploy
swa deploy ./dist --env production
```

### Any web server

Copy the `dist/` folder anywhere that serves HTTPS:

```
# Nginx example
server {
    listen 443 ssl;
    server_name addin.your-domain.com;
    root /var/www/halo-addin;
    index taskpane.html;
}
```

---

## Step 5: Upload the manifest to Microsoft 365

1. Go to **[Microsoft 365 Admin Center](https://admin.microsoft.com)**
2. Navigate to **Settings → Integrated Apps**
3. Click **Upload custom apps**
4. Choose **Upload manifest file (.xml)**
5. Select your updated `manifest.xml`
6. Click **Upload**

The add-in appears in Outlook within minutes (usually < 1 hour, sometimes up to 24).

### Verify

Open Outlook (desktop or web), select any email, and look for **"Link to Halo"** in the ribbon or the reading pane's add-in bar. Click it — the task pane should load with the banner at the top.

---

## Updating the add-in

To push an update:

1. Make your changes in `add-in/src/`
2. Rebuild: `npm run build`
3. Re-upload the `dist/` folder to your host (overwrite the existing files)
4. If you changed permissions or manifest settings, re-upload `manifest.xml` to the Admin Center

The add-in loads fresh from your host on every open — users see the update immediately, no reinstall needed.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| **Add-in doesn't appear in Outlook** | Manifest still propagating, or wrong URL | Wait 1–2 hours. Verify manifest.xml has the correct HTTPS SourceLocation. |
| **Task pane is blank / shows error** | Hosted URL doesn't match manifest, or CORS blocked | Open DevTools (F12) in Outlook on the web. Check the Console for errors. |
| **"Link to Halo" button shows but pane is stuck loading** | Office.js can't reach your hosted assets | Verify the URL is publicly accessible over HTTPS. Try `curl https://your-domain.com/taskpane.html`. |
| **Add-in works on web but not desktop** | Desktop Outlook caches the manifest longer | Clear the Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` — delete its contents and restart Outlook. |
| **manifest.xml fails validation on upload** | XML syntax error, missing elements, or non-HTTPS URLs | Validate with `npx office-addin-manifest validate manifest.xml`. Fix any errors it reports. |
| **Icons missing** | Icon URLs don't resolve | Verify each icon file exists at the URL in the manifest. Use absolute HTTPS URLs. |

---

## Security notes

- **The add-in calls Halo's API directly from the user's browser.** Halo must be reachable over HTTPS. If Halo and the user are on the same network, this works. If Halo is behind a firewall, the user's browser must have network access to it.
- **OAuth tokens are stored in `localStorage`** in the user's browser. They are scoped to the add-in's origin and expire after the configured lifetime.
- **The config.ts file is bundled into the JavaScript** at build time. Anyone who loads the add-in can see the `haloClientId` and action IDs. These are not secrets — the OAuth flow protects API access. The client secret should NOT be in `config.ts`.
- **Replace the placeholder icons.** The shipped PNGs are 1×1 transparent pixels. They work for testing but look broken to end users. Supply real square icons at all 7 sizes (16, 25, 32, 48, 64, 80, 128).