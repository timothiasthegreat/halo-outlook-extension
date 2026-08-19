/**
 * Runtime configuration for the Halo Outlook Add-in.
 *
 * On startup, fetches tenant config from the Express server's /api/config
 * endpoint (same origin — add-in is served from Express). Falls back to
 * the bundled defaults below if the fetch fails (local dev / offline).
 *
 * Config is re-fetched when the server returns a different config than
 * what's cached, or when an auth call fails with a bad client_id.
 */

export interface AddinConfig {
  /** HaloPSA instance URL, e.g. "https://your-instance.halopsa.com" */
  haloUrl: string;
  /** OAuth2 client ID for auth code flow with PKCE */
  haloClientId: string;
  /** Ticket action outcome IDs */
  actions: {
    emailReceived: number; // inbound from customer
    emailSent: number; // outbound to customer
  };
  /** Custom field ID that stores conversationId on tickets */
  customFieldConvId: number;
  /** Default ticket type for new tickets created from conversations */
    defaultTicketTypeId: number;
  }

// ── Bundled defaults (fallback when /api/config is unavailable) ──

const FALLBACK_CONFIG: AddinConfig = {
  haloUrl: "https://your-instance.halopsa.com",
  haloClientId: "",
  actions: {
    emailReceived: 0,
    emailSent: 16,
  },
  customFieldConvId: 285,
    defaultTicketTypeId: 1,
  };

// ── Runtime config cache ──────────────────────────────────────

let runtimeConfig: AddinConfig | null = null;
let loaded = false;

/**
 * Fetch tenant config from the Express server, falling back to bundled defaults.
 *
 * On first call: fetches /api/config and caches.
 * On subsequent calls: re-fetches to detect config changes. If the server
 * returns a different haloClientId or haloUrl, the cache is updated.
 * If the fetch fails, returns the existing cached config.
 */
export async function loadConfig(): Promise<AddinConfig> {
  try {
    const resp = await fetch("/api/config");
    if (resp.ok) {
      const data = (await resp.json()) as AddinConfig;
      const fresh = { ...FALLBACK_CONFIG, ...data };

      // Update cache if this is first load OR config changed
      if (!loaded ||
          runtimeConfig?.haloClientId !== fresh.haloClientId ||
          runtimeConfig?.haloUrl !== fresh.haloUrl) {
        runtimeConfig = fresh;
      }
      loaded = true;
      return runtimeConfig!;
    }
  } catch {
    // Server not reachable — use cached config if available, else fallback
  }

  if (!loaded) {
    runtimeConfig = { ...FALLBACK_CONFIG };
    loaded = true;
  }
  return runtimeConfig!;
}

/**
 * Force a config refresh from the server. Re-fetches /api/config
 * and updates the cache. Safe to call any time config may be stale.
 */

/**
 * Get the current config (must call loadConfig() first).
 * Returns the runtime config if loaded, otherwise the fallback.
 */
export function getConfig(): AddinConfig {
  return runtimeConfig ?? FALLBACK_CONFIG;
}