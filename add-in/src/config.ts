/**
 * Build-time configuration for the Halo Outlook Add-in.
 *
 * The deployer fills in these values before running `npm run build`.
 * They are bundled into the add-in and do NOT change at runtime.
 *
 * For community distribution: clone the repo, edit this file, rebuild.
 * For self-hosted config: point `configUrl` at a JSON endpoint instead.
 */
export interface AddinConfig {
  /** HaloPSA instance URL, e.g. "https://your-instance.halopsa.com" */
  haloUrl: string;
  /** OAuth2 client ID for auth code flow with PKCE */
  haloClientId: string;
  /** Ticket action outcome IDs */
  actions: {
    emailReceived: number; // inbound from customer (typically 0)
    emailSent: number; // outbound to customer (typically 16)
  };
  /** Custom field ID that stores conversationId on tickets */
  customFieldConvId: number;
  /** Default ticket type for new tickets created from conversations */
  defaultTicketTypeId: number;
}

/**
 * DEFAULT CONFIG — Edit these values for your Halo instance.
 *
 * Every deployer MUST change haloUrl, haloClientId, and verify
 * the action IDs match their Halo instance's configured actions.
 */
const config: AddinConfig = {
  haloUrl: "https://your-instance.halopsa.com",
  haloClientId: "",
  actions: {
    emailReceived: 0, // "Email Update" / inbound
    emailSent: 16, // "Email User" / outbound
  },
  customFieldConvId: 285, // CFticketconvid — discover yours with setup_check.py
  defaultTicketTypeId: 1, // Incident — adjust per your instance
};

export default config;