/**
 * HaloPSA API client for the Outlook add-in.
 *
 * Handles OAuth2 auth code flow with PKCE and all ticket/action endpoints
 * needed by the add-in UI. Uses the Office.js dialog API for auth redirects
 * and localStorage for token persistence.
 *
 * Config is fetched dynamically from /api/config at startup (see config.ts).
 * All Halo URL / client ID references use getConfig() so they reflect the
 * current tenant config without requiring a rebuild.
 */

import { getConfig, loadConfig } from "./config";

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
 * Returns after the user authenticates and the auth code is captured.
 *
 * The flow:
 *   1. Open Office dialog to Halo's /auth endpoint
 *   2. User logs in and authorizes the app
 *   3. Halo redirects to auth-complete.html?code=...
 *   4. auth-complete.html calls messageParent with the code
 *   5. We receive the code, exchange it for a token, and store it
 */
async function authorizeViaDialog(): Promise<string> {
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);
  const cfg = getConfig();

  const authUrl =
    `${cfg.haloUrl}/auth/authorize?` +
    `response_type=code&` +
    `client_id=${encodeURIComponent(cfg.haloClientId)}&` +
    `redirect_uri=${encodeURIComponent(getRedirectUri())}&` +
    `code_challenge=${codeChallenge}&` +
    `code_challenge_method=S256&` +
    `scope=all`;

  return new Promise<string>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 40, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(`Dialog failed: ${result.error.message}`));
          return;
        }
        const dialog = result.value;

        // Receive the auth code from auth-complete.html
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          async (arg) => {
            dialog.close();
            try {
              // arg can be { message, origin } on success or { error } on dialog error
              if ("error" in arg) {
                reject(new Error(`Dialog error: ${arg.error}`));
                return;
              }
              const msg = JSON.parse(arg.message);
              if (msg.error) {
                reject(new Error(`Authorization failed: ${msg.error}`));
                return;
              }
              if (!msg.code) {
                reject(new Error("No authorization code received"));
                return;
              }

              // Exchange the auth code for tokens
              const tokenData = await exchangeCodeForToken(msg.code, codeVerifier);
              storeToken(tokenData);
              resolve(tokenData.access_token);
            } catch (err: any) {
              reject(err);
            }
          }
        );

        // Handles: user closes dialog, or Office error event
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => {
            dialog.close();
            reject(new Error("Login was cancelled or the dialog was closed"));
          }
        );
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
  const cfg = getConfig();
  const resp = await fetch(`${cfg.haloUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      client_id: cfg.haloClientId,
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
  const cfg = getConfig();
  const resp = await fetch(`${cfg.haloUrl}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      client_id: cfg.haloClientId,
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
  const cfg = getConfig();
  const url = `${cfg.haloUrl}/api${path}`;
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
  } catch (err: any) {
    // If we got here with a token that failed, the config (client_id) may
    // have changed. Re-fetch from the server before launching the dialog.
    // loadConfig() always hits the server (no short-circuit on subsequent calls
    // in config.ts v2), so this picks up any config.yaml changes.
    if (err.message?.includes("Token refresh failed") ||
        err.message?.includes("Not authenticated")) {
      await loadConfig();
    }
    // No valid token — launch the OAuth dialog
    // authorizeViaDialog() handles the full PKCE flow and stores the token
    await authorizeViaDialog();
    // Verify we now have a working token
    await getAccessToken();
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
  const cfg = getConfig();
  const payload = [
    {
      summary: params.summary.slice(0, 70),
      details_html: params.details_html,
      tickettype_id: params.tickettype_id ?? cfg.defaultTicketTypeId,
      user_id: params.user_id,
      customfields: [
        {
          id: cfg.customFieldConvId,
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
  const cfg = getConfig();
  const payload = [
    {
      id: ticketId,
      customfields: [
        {
          id: cfg.customFieldConvId,
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
  return [];
}

// ── Office.js context helpers ───────────────────────────────────

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
      resolve("");
    }
  });
}

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

export function getCurrentBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox?.item;
    if (!item) {
      reject(new Error("No mailbox item available"));
      return;
    }
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