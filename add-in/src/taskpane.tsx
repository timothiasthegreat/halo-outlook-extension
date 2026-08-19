/**
 * Main task pane UI for the Halo Outlook Add-in.
 *
 * Always shows one of two states:
 *   Untracked: "Create Ticket" + "Link to Ticket" buttons
 *   Tracked:   Ticket info banner + recent activity + "Unlink" button
 *
 * Initializes Office.js, loads the current email context, and
 * checks whether the conversation is already tracked.
 */

import * as React from "react";
import * as ReactDOM from "react-dom";

import Banner from "./banner";
import * as halo from "./halo";
import { loadConfig, getConfig } from "./config";
import { Setup } from "./setup";

// ── state ──────────────────────────────────────────────────────

interface TaskpaneState {
  // Ticket tracking state
  ticket: halo.HaloTicket | null;
  loading: boolean;
  error: string | null;
  lastSyncAt: string | null;

  // Conversation context
  conversationId: string;
  subject: string;
  senderEmail: string;
  userEmail: string;

  // UI state
  mode: "untracked" | "tracked" | "link" | "login" | "setup";
  ticketTypeId: number;
  ticketTypes: Array<{ id: number; name: string }>;
  ticketTypesLoading: boolean;
  searchQuery: string;
  searchResults: halo.HaloTicket[];
  searching: boolean;
  creating: boolean;
  linking: boolean;
}

// ── styles ─────────────────────────────────────────────────────

const paneStyle: React.CSSProperties = {
  padding: "12px",
  fontFamily: "Segoe UI, sans-serif",
  fontSize: "13px",
  color: "#333",
  maxWidth: "100%",
  boxSizing: "border-box",
};

const buttonStyle: React.CSSProperties = {
  padding: "6px 16px",
  border: "1px solid #0078d4",
  borderRadius: "2px",
  background: "#0078d4",
  color: "white",
  cursor: "pointer",
  fontSize: "13px",
  fontWeight: 600,
  marginRight: "8px",
  marginBottom: "8px",
};

const secondaryButtonStyle: React.CSSProperties = {
  ...buttonStyle,
  background: "white",
  color: "#0078d4",
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: "6px 8px",
  border: "1px solid #ccc",
  borderRadius: "2px",
  fontSize: "13px",
  boxSizing: "border-box",
  marginBottom: "8px",
};

const selectStyle: React.CSSProperties = {
  ...inputStyle,
  width: "100%",
};

// ── component ──────────────────────────────────────────────────

class Taskpane extends React.Component<{}, TaskpaneState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      ticket: null,
      loading: true,
      error: null,
      lastSyncAt: null,
      conversationId: "",
      subject: "",
      senderEmail: "",
      userEmail: "",
      mode: "untracked",
      ticketTypeId: getConfig().defaultTicketTypeId,
      ticketTypes: [],
      ticketTypesLoading: false,
      searchQuery: "",
      searchResults: [],
      searching: false,
      creating: false,
      linking: false,
    };
  }

  async componentDidMount(): Promise<void> {
    // Load runtime config from Express server first, then initialize
    Office.onReady()
      .then(() => loadConfig())
      .then(() => {
        // Listen for email selection changes — when the taskpane is
        // pinned open, componentDidMount only fires once. ItemChanged
        // reloads context each time the user clicks a different email.
        if (Office.context.mailbox) {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            (_eventArgs: any) => this.handleItemChanged()
          );
        }
      })
      .then(() => this.loadContext())
      .then(() => this.checkAuth())
      .then(() => this.fetchTicketTypes())
      .catch((err) => this.setState({ error: err.message, loading: false }));
  }

  componentWillUnmount(): void {
    // Clean up ItemChanged handler when the taskpane is closed
    try {
      if (Office.context.mailbox) {
        Office.context.mailbox.removeHandlerAsync(
          Office.EventType.ItemChanged,
          (_result: any) => {} // fire-and-forget; pane is closing
        );
      }
    } catch {
      // Office.js may not be available during teardown
    }
  }

  /**
   * Called when the user clicks a different email while the taskpane
   * is pinned open. Resets state and reloads context for the new item.
   */
  handleItemChanged(): void {
    // Reset to loading/untracked until we know the new item's state
    this.setState({
      ticket: null,
      loading: true,
      error: null,
      lastSyncAt: null,
      conversationId: "",
      subject: "",
      senderEmail: "",
      userEmail: "",
      mode: "untracked",
      searchQuery: "",
      searchResults: [],
      searching: false,
      creating: false,
      linking: false,
    });
    this.loadContext();
  }

  /**
   * Fetch ticket types from the Express server (which proxies to Halo).
   * Sends the user's OAuth token so results respect agent permissions.
   * Falls back to a single default type if the fetch fails.
   */
  async fetchTicketTypes(): Promise<void> {
    const defaultTypeId = getConfig().defaultTicketTypeId;
    this.setState({ ticketTypesLoading: true });
    try {
      const token = localStorage.getItem("halo_outlook_token");
      const tokenData = token ? JSON.parse(token) : null;
      if (!tokenData?.access_token) {
        throw new Error("Not authenticated");
      }
      const resp = await fetch("/api/ticket-types", {
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
        },
      });
      if (!resp.ok) throw new Error(`Server returned ${resp.status}`);
      const types = await resp.json();
      if (types.length > 0) {
        this.setState({ ticketTypes: types, ticketTypesLoading: false });
        // If current ticketTypeId isn't in the list, switch to default
        if (!types.find((t: any) => t.id === this.state.ticketTypeId)) {
          this.setState({ ticketTypeId: defaultTypeId });
        }
        return;
      }
    } catch {
      // Fall back to a single default type
    }
    this.setState({
      ticketTypes: [{ id: defaultTypeId, name: "Ticket" }],
      ticketTypesLoading: false,
    });
  }

  /**
   * Check whether we have a valid Halo token. If not, switch to login mode.
   * Only checks localStorage — does NOT trigger the auth dialog.
   */
  async checkAuth(): Promise<void> {
    // Check localStorage directly for a stored token (no dialog trigger)
    const raw = localStorage.getItem("halo_outlook_token");
    if (!raw) {
      this.setState({ mode: "login" });
      return;
    }
    try {
      const token = JSON.parse(raw);
      // If expired, show login
      if (!token.access_token || Date.now() > token.expires_at - 60000) {
        this.setState({ mode: "login" });
      }
      // Otherwise, user is authenticated — stay in current mode
    } catch {
      this.setState({ mode: "login" });
    }
  }

  async loadContext(): Promise<void> {
    try {
      const [convId, subject, senderEmail, userEmail] = await Promise.all([
        halo.getCurrentConversationId(),
        halo.getCurrentSubject(),
        halo.getCurrentSenderEmail(),
        halo.getCurrentUserEmail(),
      ]);
      this.setState({ conversationId: convId, subject, senderEmail, userEmail });

      if (!convId) {
        this.setState({ loading: false });
        return;
      }

      // Check if this conversation is already tracked.
      // The primary mechanism is the watcher's state store, but
      // the add-in can also check via localStorage cache or a
      // Halo API query for tickets with this conversationId.
      //
      // For MVP: look up in localStorage cache set by previous
      // "Create Ticket" actions in this add-in session.
      const cached = this.getCachedTicket(convId);
      if (cached) {
        this.setState({
          ticket: cached,
          mode: "tracked",
          loading: false,
          lastSyncAt: this.getCachedLastSync(convId),
        });
        return;
      }

      this.setState({ loading: false });
    } catch (err: any) {
      this.setState({ error: err.message, loading: false });
    }
  }

  // ── localStorage helpers ───────────────────────────────────

  getCachedTicket(conversationId: string): halo.HaloTicket | null {
    try {
      const raw = localStorage.getItem(`halo_ticket_${conversationId}`);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  setCachedTicket(conversationId: string, ticket: halo.HaloTicket): void {
    localStorage.setItem(
      `halo_ticket_${conversationId}`,
      JSON.stringify(ticket)
    );
    localStorage.setItem(
      `halo_last_sync_${conversationId}`,
      new Date().toISOString()
    );
  }

  getCachedLastSync(conversationId: string): string | null {
    return localStorage.getItem(`halo_last_sync_${conversationId}`);
  }

  /**
   * Register the user's mailbox with the watcher so future replies
   * in this conversation are picked up by Graph polling.
   * Uses the refresh token from the OAuth flow. Non-blocking — failure
   * is silent (the watcher picks up missed messages on its next poll).
   */
  async registerMailbox(email: string): Promise<void> {
      if (!email) return;
      try {
        const tokenRaw = localStorage.getItem("halo_outlook_token");
        if (!tokenRaw) {
          console.warn("[registerMailbox] No token in localStorage");
          return;
        }
        const tokenData = JSON.parse(tokenRaw);

        // Use refresh_token if available, otherwise pass the access_token
        // (the watcher can use it for initial Graph auth, then refresh on its own)
        const refreshToken = tokenData?.refresh_token || tokenData?.access_token;
        if (!refreshToken) {
          console.warn("[registerMailbox] No token available to register");
          return;
        }

        const resp = await fetch("/api/register", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            email,
            refresh_token: refreshToken,
          }),
        });
        if (!resp.ok) {
          console.warn(`[registerMailbox] Registration failed: ${resp.status}`);
        } else {
          console.log(`[registerMailbox] Registered ${email} successfully`);
        }
      } catch (err: any) {
        console.warn("[registerMailbox] Error:", err.message);
      }
    }

  /**
   * Register this conversation in the watcher's state.db so the sync engine
   * picks it up. Non-blocking — failure is silent (the add-in still shows
   * the ticket as linked via localStorage, and the watcher retries on its
   * next poll cycle).
   */
  async trackConversation(conversationId: string, ticketId: number): Promise<void> {
      if (!conversationId || !ticketId) return;
      try {
        const resp = await fetch("/api/conversations", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            conversationId,
            ticketId,
            watchedBy: this.state.userEmail || undefined,
          }),
        });
        if (!resp.ok) {
          console.warn(`[trackConversation] Failed: ${resp.status}`);
        } else {
          console.log(`[trackConversation] Tracked ${conversationId} → ticket ${ticketId}`);
        }
      } catch (err: any) {
        console.warn("[trackConversation] Error:", err.message);
      }
    }

  // ── actions ─────────────────────────────────────────────────

  handleCreateTicket = async (): Promise<void> => {
    const { conversationId, subject, senderEmail, userEmail, ticketTypeId } = this.state;
    if (!conversationId) {
      this.setState({ error: "No conversation selected" });
      return;
    }

    this.setState({ creating: true, error: null });
    try {
      const body = await halo.getCurrentBody();

      // Look up the user in Halo by email so the ticket has an owner
      const userId = await halo.findUserByEmail(senderEmail);

      const ticket = await halo.createTicket({
        summary: subject || "Email from Outlook",
        details_html: body || `<p>Email from ${senderEmail}</p>`,
        tickettype_id: ticketTypeId,
        user_id: userId,
        conversationId,
      });

      // Halo's POST /Tickets response may not include all fields.
      // Re-fetch the full ticket to get ticketnumber, status_name, etc.
      let fullTicket = ticket;
      try {
        fullTicket = await halo.getTicket(ticket.id);
      } catch {
        // Use whatever we got from the create response
      }

      // Cache the ticket for this session
      this.setCachedTicket(conversationId, fullTicket);
      this.setState({
        ticket: fullTicket,
        mode: "tracked",
        creating: false,
        lastSyncAt: new Date().toISOString(),
      });

      // Register the user's mailbox with the watcher so future replies
      // are picked up by the Graph polling loop. Uses the refresh token
      // from the OAuth flow stored in localStorage.
      this.registerMailbox(userEmail);

      // Also journal the initial email as an action
      const internetMessageId = await halo.getCurrentInternetMessageId();
      if (internetMessageId) {
        try {
          await halo.createAction({
            ticket_id: fullTicket.id,
            outcome_id: getConfig().actions.emailReceived,
            note: `Subject: ${subject}\nFrom: ${senderEmail}\n\n${stripHtml(body)}`,
            note_html: body || `<p>Email from ${senderEmail}</p>`,
            email_message_id: internetMessageId,
          });
        } catch {
          // Journal failure is non-fatal — the watcher will pick it up
        }
      }
    } catch (err: any) {
      // If auth error, switch to login mode so user can sign in
      if (err.message?.includes("Not authenticated") || err.message?.includes("Session expired")) {
        this.setState({ mode: "login", creating: false });
      } else {
        this.setState({ error: err.message, creating: false });
      }
    }
  };

  handleSearchTickets = async (): Promise<void> => {
    const { searchQuery } = this.state;
    if (!searchQuery.trim()) return;

    this.setState({ searching: true });
    try {
      const tickets = await halo.searchTickets(searchQuery.trim());
      this.setState({ searchResults: tickets, searching: false });
    } catch (err: any) {
      if (err.message?.includes("Not authenticated") || err.message?.includes("Session expired")) {
        this.setState({ mode: "login", searching: false });
      } else {
        this.setState({ error: err.message, searching: false });
      }
    }
  };

  handleLinkTicket = async (ticketId: number): Promise<void> => {
    const { conversationId, subject, userEmail } = this.state;
    if (!conversationId) {
      this.setState({ error: "No conversation selected" });
      return;
    }

    this.setState({ linking: true, error: null });
    try {
      await halo.linkConversation(ticketId, conversationId);
      const ticket = await halo.getTicket(ticketId);

      this.setCachedTicket(conversationId, ticket);
      this.setState({
        ticket,
        mode: "tracked",
        linking: false,
        lastSyncAt: new Date().toISOString(),
      });

      // Register the user's mailbox with the watcher so future replies
      // are picked up by the Graph polling loop. Uses the refresh token
      // from the OAuth flow stored in localStorage.
      this.registerMailbox(userEmail);

      // Track the conversation in the watcher's state.db so the sync
      // engine knows which conversations to poll.
      this.trackConversation(conversationId, ticketId);
    } catch (err: any) {
      if (err.message?.includes("Not authenticated") || err.message?.includes("Session expired")) {
        this.setState({ mode: "login", linking: false });
      } else {
        this.setState({ error: err.message, linking: false });
      }
    }
  };

  handleUnlink = async (): Promise<void> => {
    const { conversationId } = this.state;
    if (!conversationId) return;

    localStorage.removeItem(`halo_ticket_${conversationId}`);
    localStorage.removeItem(`halo_last_sync_${conversationId}`);
    this.setState({ ticket: null, mode: "untracked", lastSyncAt: null });
  };

  handleLogin = (): void => {
    halo
      .ensureAuthenticated()
      .then(() => {
        // Auth succeeded — go back to untracked mode with fresh context
        this.setState({ mode: "untracked", error: null });
        this.loadContext();
        this.fetchTicketTypes();
      })
      .catch((err) => this.setState({ error: err.message }));
  };

  // ── render ──────────────────────────────────────────────────

  render(): React.ReactNode {
    const {
      ticket,
      loading,
      error,
      mode,
      searchResults,
      searching,
      creating,
      linking,
      searchQuery,
      subject,
      ticketTypeId,
      ticketTypes,
    } = this.state;

    return (
      <div style={paneStyle}>
        {/* ── Header ── */}
        <div
          style={{
            fontSize: "14px",
            fontWeight: 700,
            marginBottom: "12px",
            display: "flex",
            alignItems: "center",
            gap: "6px",
          }}
        >
          🔗 Halo Ticket Bridge
        </div>

        {/* ── Banner — always visible ── */}
        <Banner
          ticket={ticket}
          loading={loading}
          lastSyncAt={this.state.lastSyncAt}
        />

        {/* ── Error ── */}
        {error && (
          <div
            style={{
              color: "#a80000",
              background: "#fde7e9",
              padding: "8px",
              borderRadius: "4px",
              marginBottom: "12px",
              fontSize: "12px",
            }}
          >
            {error}
            <button
              style={{
                marginLeft: "8px",
                background: "none",
                border: "none",
                cursor: "pointer",
                color: "#a80000",
                textDecoration: "underline",
              }}
              onClick={() => this.setState({ error: null })}
            >
              Dismiss
            </button>
          </div>
        )}

        {/* ── Untracked mode ── */}
        {mode === "untracked" && !loading && (
          <div>
            <div
              style={{
                color: "#666",
                marginBottom: "8px",
                fontSize: "12px",
              }}
            >
              Select a ticket type and create a ticket from this conversation,
              or link it to an existing Halo ticket.
            </div>

            <label style={{ fontSize: "12px", fontWeight: 600 }}>Ticket type:</label>
            <select
              style={selectStyle}
              value={ticketTypeId}
              onChange={(e) => this.setState({ ticketTypeId: Number(e.target.value) })}
            >
              {ticketTypes.map((tt) => (
                <option key={tt.id} value={tt.id}>{tt.name}</option>
              ))}
            </select>

            <button
              style={buttonStyle}
              onClick={this.handleCreateTicket}
              disabled={creating}
            >
              {creating ? "Creating…" : "Create Ticket"}
            </button>
            <button
              style={secondaryButtonStyle}
              onClick={() => this.setState({ mode: "link" })}
            >
              Link to Ticket…
            </button>
          </div>
        )}

        {/* ── Link mode ── */}
        {mode === "link" && (
          <div>
            <div style={{ marginBottom: "8px", fontSize: "12px", color: "#666" }}>
              Search for a ticket by reference number (e.g. TKT-1234) or
              customer name, then select it to link this conversation.
            </div>
            <input
              style={inputStyle}
              placeholder="Search tickets…"
              value={searchQuery}
              onChange={(e) => this.setState({ searchQuery: e.target.value })}
              onKeyDown={(e) => e.key === "Enter" && this.handleSearchTickets()}
            />
            <button
              style={buttonStyle}
              onClick={this.handleSearchTickets}
              disabled={searching || !searchQuery.trim()}
            >
              {searching ? "Searching…" : "Search"}
            </button>
            <button
              style={secondaryButtonStyle}
              onClick={() => this.setState({ mode: "untracked" })}
            >
              Back
            </button>

            {searchResults.length > 0 && (
              <div style={{ marginTop: "8px" }}>
                {searchResults.map((t) => (
                  <div
                    key={t.id}
                    style={{
                      border: "1px solid #e0e0e0",
                      borderRadius: "4px",
                      padding: "8px",
                      marginBottom: "6px",
                      cursor: "pointer",
                    }}
                    onClick={() => this.handleLinkTicket(t.id)}
                  >
                    <div style={{ fontWeight: 600 }}>
                      {t.ticketnumber ? `#${t.ticketnumber}` : `#${t.id}`} —{" "}
                      {t.summary.slice(0, 60)}
                    </div>
                    <div style={{ fontSize: "12px", color: "#666" }}>
                      {t.status_name ?? `Status ${t.status_id}`}
                      {t.assigned_to_name && ` • ${t.assigned_to_name}`}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ── Tracked mode — additional controls ── */}
        {mode === "tracked" && ticket && (
          <div style={{ marginTop: "8px" }}>
            <div
              style={{
                fontSize: "12px",
                color: "#999",
                marginBottom: "8px",
              }}
            >
              New replies in this conversation will automatically journal to
              ticket {ticket.ticketnumber ?? `#${ticket.id}`}.
            </div>
            <button
              style={{
                ...secondaryButtonStyle,
                color: "#a80000",
                borderColor: "#a80000",
              }}
              onClick={this.handleUnlink}
            >
              Unlink
            </button>
          </div>
        )}

        {/* ── Login prompt ── */}
                {mode === "login" && (
                  <div style={{ textAlign: "center", padding: "20px" }}>
                    <div style={{ marginBottom: "12px", color: "#666" }}>
                      Connect to HaloPSA to start tracking conversations.
                    </div>
                    <button style={buttonStyle} onClick={this.handleLogin}>
                      Sign in to Halo
                    </button>
                  </div>
                )}

                {/* ── Setup prompt ── */}
                {mode === "setup" && (
                  <Setup
                    email={this.state.senderEmail}
                    refreshToken=""
                    onRegistered={() => {
                      localStorage.setItem("halo_watcher_registered", "true");
                      this.setState({ mode: "untracked" });
                    }}
                    onSkip={() => this.setState({ mode: "untracked" })}
                  />
                )}
      </div>
    );
  }
}

// ── mount ──────────────────────────────────────────────────────

Office.onReady(() => {
  ReactDOM.render(React.createElement(Taskpane), document.getElementById("root"));
});

// ── helper ─────────────────────────────────────────────────────

function stripHtml(html: string): string {
  if (typeof document !== "undefined") {
    const div = document.createElement("div");
    div.innerHTML = html;
    return div.textContent ?? div.innerText ?? "";
  }
  return html.replace(/<[^>]*>/g, "");
}