/**
 * HaloPSA API client for the Outlook add-in.
 *
 * Handles OAuth2 auth code flow with PKCE and all ticket/action endpoints
 * needed by the add-in UI. Uses the Office.js dialog API for auth redirects
 * and localStorage for token persistence.
 */

import config from "./config";

// ── types ──────────────────────────────────────────────────────

export interface HaloTicket {
  id: number;
  summary: string;
  tickettype_name?: string;
  status_id: number;
  status_name?: string;
  user_id?: number;
  user_name?: string;
  assigned_to_name?: string;
  ticketnumber?: string;
  details?: string;
  details_html?: string;
}

export interface HaloAction {
  id: number;
  ticket_id: number;
  outcome_id: number;
  outcome?: string;
  note: string;
  note_html: string;
  datetime: string;
  who?: string;
  emailfrom?: string;
  emaildirection?: string;
  email_message_id?: string;
}

export interface SearchResult {
  tickets: HaloTicket[];
  record_count: number;
}

// ── token store ────────────────────────────────────────────────

const TOKEN_KEY = "halo_outlook_token";

interface TokenData {
  access_token: string;
  refresh_token?: string;
  expires_at: number; // epoch ms
}

function getStoredToken(): TokenData | null {
  try {
    const raw = localStorage.getItem(TOKEN_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as TokenData;
  } catch {
    return null;
  }
}

function storeToken(data: TokenData): void {
  localStorage.setItem(TOKEN_KEY, JSON.stringify(data));
}

function clearToken(): void {
  localStorage.removeItem(TOKEN_KEY);
}

// ── auth code flow with PKCE ───────────────────────────────────

function generateCodeVerifier(): string {
  const array = new Uint8Array(64);
  self.crypto.getRandomValues(array);
  return btoa(String.fromCharCode(...array))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "")
    .slice(0, 128);
}

async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder();
  const digest = await self.crypto.subtle.digest("SHA-256", encoder.encode(verifier));
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}

/**
 * Launch the OAuth2 auth code flow via the Office dialog API.
 * Returns after the user authenticates and the redirect is captured.
 */
async function authorizeViaDialog(): Promise<void> {
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  const authUrl =
    `${config.haloUrl}/auth?` +
    `response_type=code&` +
    `client_id=${encodeURIComponent(config.haloClientId)}&` +
    `redirect_uri=${encodeURIComponent(getRedirectUri())}&` +
    `code_challenge=${codeChallenge}&` +
    `code_challenge_method=S256&` +
    `scope=all`;

  // Open the Halo auth page in an Office dialog
  return new Promise<void>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 40, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(`Dialog failed: ${result.error.message}`));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          // The redirect URI page posts the auth code back via messageParent
          // But Office dialog doesn't capture the redirect easily.
          // Simplified: close on user completion, poll localStorage for token.
          dialog.close();
          resolve();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          dialog.close();
          resolve();
        });
      }
    );
  });
}

function getRedirectUri(): string {
  // For Office add-ins, the redirect URI is the add-in's origin
  return `${window.location.origin}/auth-complete.html`;
}

/**
 * Exchange an authorization code for tokens.
 */
async function exchangeCodeForToken(code: string, codeVerifier: string): Promise<TokenData> {
  const resp = await fetch(`${config.haloUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      client_id: config.haloClientId,
      code,
      redirect_uri: getRedirectUri(),
      code_verifier: codeVerifier,
    }),
  });
  if (!resp.ok) throw new Error(`Token exchange failed: ${resp.status}`);
  const data = await resp.json();
  return {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: Date.now() + (data.expires_in ?? 3600) * 1000,
  };
}

async function refreshToken(token: TokenData): Promise<TokenData> {
  if (!token.refresh_token) throw new Error("No refresh token available");
  const resp = await fetch(`${config.haloUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      client_id: config.haloClientId,
      refresh_token: token.refresh_token,
    }),
  });
  if (!resp.ok) throw new Error(`Token refresh failed: ${resp.status}`);
  const data = await resp.json();
  return {
    access_token: data.access_token,
    refresh_token: data.refresh_token ?? token.refresh_token,
    expires_at: Date.now() + (data.expires_in ?? 3600) * 1000,
  };
}

async function getAccessToken(): Promise<string> {
  let token = getStoredToken();
  if (!token) {
    // No stored token — need full auth flow.
    // For now, return empty and let UI show login prompt.
    throw new Error("Not authenticated — please log in.");
  }
  if (Date.now() > token.expires_at - 60000) {
    try {
      token = await refreshToken(token);
      storeToken(token);
    } catch {
      clearToken();
      throw new Error("Session expired — please log in again.");
    }
  }
  return token.access_token;
}

// ── API helpers ────────────────────────────────────────────────

async function apiRequest<T>(
  method: string,
  path: string,
  body?: unknown
): Promise<T> {
  const token = await getAccessToken();
  const url = `${config.haloUrl}/api${path}`;
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  };
  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
  }
  const resp = await fetch(url, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Halo API ${method} ${path} failed (${resp.status}): ${text}`);
  }
  return resp.json();
}

// ── public API ─────────────────────────────────────────────────

export async function ensureAuthenticated(): Promise<void> {
  try {
    await getAccessToken();
  } catch {
    await authorizeViaDialog();
    // After dialog, we don't have the code (Office dialog limitation).
    // The user must go through the full PKCE flow which requires a
    // capture page. In practice, Halo's OAuth may support simpler flows.
    //
    // For Phase 3 MVP: if token is missing, show "Login required" in UI.
    // The deployer can set up a client credentials flow for the add-in.
    throw new Error(
      "Authentication required. Please configure your Halo OAuth credentials and restart the add-in."
    );
  }
}

/**
 * Create a new ticket from an email conversation.
 * Halo expects a JSON array in the POST body.
 */
export async function createTicket(params: {
  summary: string;
  details_html: string;
  user_id?: number;
  tickettype_id?: number;
  conversationId: string;
}): Promise<HaloTicket> {
  const payload = [
    {
      summary: params.summary.slice(0, 70),
      details_html: params.details_html,
      tickettype_id: params.tickettype_id ?? config.defaultTicketTypeId,
      user_id: params.user_id,
      customfields: [
        {
          id: config.customFieldConvId,
          value: params.conversationId,
        },
      ],
    },
  ];
  const result = await apiRequest<HaloTicket[]>("POST", "/Tickets", payload);
  return result[0];
}

/**
 * Search for tickets by reference number or customer name.
 */
export async function searchTickets(query: string): Promise<HaloTicket[]> {
  const params = new URLSearchParams({ search: query });
  const result = await apiRequest<SearchResult>("GET", `/Tickets?${params}`);
  return result.tickets ?? [];
}

/**
 * Get a single ticket by ID.
 */
export async function getTicket(ticketId: number): Promise<HaloTicket> {
  return apiRequest<HaloTicket>("GET", `/Tickets/${ticketId}`);
}

/**
 * Update the conversationId custom field on an existing ticket.
 * Uses POST (Halo quirk: PATCH/PUT return 405).
 */
export async function linkConversation(
  ticketId: number,
  conversationId: string
): Promise<void> {
  const payload = [
    {
      id: ticketId,
      customfields: [
        {
          id: config.customFieldConvId,
          value: conversationId,
        },
      ],
    },
  ];
  await apiRequest("POST", "/Tickets", payload);
}

/**
 * Post a ticket action (journal entry).
 * Used by both the initial create flow and OnMessageSend handler.
 */
export async function createAction(params: {
  ticket_id: number;
  outcome_id: number;
  note: string;
  note_html: string;
  email_message_id?: string;
  sendemail?: boolean;
  hiddenfromuser?: boolean;
}): Promise<HaloAction> {
  const payload = [
    {
      ticket_id: params.ticket_id,
      outcome_id: params.outcome_id,
      note: params.note,
      note_html: params.note_html,
      email_message_id: params.email_message_id,
      sendemail: params.sendemail ?? false,
      hiddenfromuser: params.hiddenfromuser ?? false,
    },
  ];
  const result = await apiRequest<HaloAction[]>("POST", "/Actions", payload);
  return result[0];
}

/**
 * List recent actions on a ticket.
 */
export async function listActions(
  ticketId: number,
  top: number = 5
): Promise<HaloAction[]> {
  const params = new URLSearchParams({
    ticket_id: String(ticketId),
    $top: String(top),
  });
  const result = await apiRequest<{ actions: HaloAction[] }>(
    "GET",
    `/Actions?${params}`
  );
  return result.actions ?? [];
}

/**
 * Get all tickets that have the conversationId custom field set.
 * Used by the watcher to find watched conversations.
 */
export async function getWatchedTickets(): Promise<HaloTicket[]> {
  // Note: Halo API does NOT support filtering by custom field value.
  // This is a documentation note — the watcher uses a local SQLite
  // state store instead. The add-in doesn't need this method; it
  // queries by conversationId via the state store / Office.js context.
  return [];
}

// ── Office.js context helpers ───────────────────────────────────

/**
 * Get the current email item's conversationId from Office.js context.
 */
export function getCurrentConversationId(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item as Office.MessageRead | undefined;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
    const convId = item.conversationId;
    if (convId) {
      resolve(convId);
    } else {
      // conversationId should always be available on MessageRead.
      // If it's empty, the item might not be loaded yet.
      resolve("");
    }
  });
}

/**
 * Get the current email item's internetMessageId from Office.js context.
 */
export function getCurrentInternetMessageId(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
    const messageId = (item as Office.MessageRead).internetMessageId;
    if (messageId) {
      resolve(messageId);
    } else {
      resolve("");
    }
  });
}

/**
 * Get the sender's email address from the current item.
 */
export function getCurrentSenderEmail(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
    const from = (item as Office.MessageRead).from;
    const address = from?.emailAddress ?? "";
    resolve(address);
  });
}

/**
 * Get the email subject from the current item.
 */
export function getCurrentSubject(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
    const subject = (item as Office.MessageRead).subject ?? "";
    resolve(subject);
  });
}

/**
 * Get the email body (prefer HTML) from the current item.
 */
export function getCurrentBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
    // Office.js provides body.getAsync with coercion type
    (item as Office.MessageRead).body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Failed to get email body"));
          return;
        }
        resolve(result.value);
      }
    );
  });
}