/**
 * Runtime configuration for the Halo Outlook Add-in.
 *
 * On startup, fetches tenant config from the Express server's /api/config
 * endpoint (same origin — add-in is served from Express). Falls back to
 * the bundled defaults below if the fetch fails (local dev / offline).
 *
 * All consumers call `await getConfig()` instead of importing directly.
 * This keeps the module-level API simple while supporting dynamic updates.
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
 * Safe to call multiple times — fetches once and caches.
 */
export async function loadConfig(): Promise<AddinConfig> {
  if (loaded) return getConfig();

  try {
    const resp = await fetch("/api/config");
    if (resp.ok) {
      const data = (await resp.json()) as AddinConfig;
      runtimeConfig = { ...FALLBACK_CONFIG, ...data };
      loaded = true;
      return runtimeConfig;
    }
  } catch {
    // Server not reachable — use fallback (local dev, offline)
  }

  runtimeConfig = { ...FALLBACK_CONFIG };
  loaded = true;
  return runtimeConfig;
}

/**
 * Get the current config (must call loadConfig() first).
 * Returns the runtime config if loaded, otherwise the fallback.
 */
export function getConfig(): AddinConfig {
  return runtimeConfig ?? FALLBACK_CONFIG;
}

// For backward compatibility — direct import still works, but consumers
// should migrate to `await getConfig()`.
const config: AddinConfig = FALLBACK_CONFIG;
export default config;