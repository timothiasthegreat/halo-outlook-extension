/**
 * Setup UI component for first-time registration.
 *
 * Prompts the user to enable email watching and optionally add
 * shared mailboxes. Posts the PKCE refresh token to /api/register
 * so the watcher can sync on their behalf.
 */

import * as React from "react";

export interface SetupProps {
  /** The user's primary email (auto-detected from Outlook) */
  email: string;
  /** The PKCE refresh token from Halo OAuth */
  refreshToken: string;
  /** Called when registration completes successfully */
  onRegistered: (additionalEmails: string[]) => void;
  /** Called when user wants to skip setup */
  onSkip: () => void;
}

interface SetupState {
  additionalEmails: string[];
  newEmail: string;
  submitting: boolean;
  error: string | null;
}

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: "6px 8px",
  border: "1px solid #ccc",
  borderRadius: "2px",
  fontSize: "13px",
  boxSizing: "border-box",
  marginBottom: "6px",
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
};

const secondaryButtonStyle: React.CSSProperties = {
  ...buttonStyle,
  background: "white",
  color: "#0078d4",
};

export class Setup extends React.Component<SetupProps, SetupState> {
  constructor(props: SetupProps) {
    super(props);
    this.state = {
      additionalEmails: [],
      newEmail: "",
      submitting: false,
      error: null,
    };
  }

  handleAddEmail = (): void => {
    const { newEmail, additionalEmails } = this.state;
    const trimmed = newEmail.trim().toLowerCase();
    if (!trimmed || !trimmed.includes("@")) return;
    if (additionalEmails.includes(trimmed)) return;

    this.setState({
      additionalEmails: [...additionalEmails, trimmed],
      newEmail: "",
    });
  };

  handleRemoveEmail = (email: string): void => {
    this.setState({
      additionalEmails: this.state.additionalEmails.filter((e) => e !== email),
    });
  };

  handleSubmit = async (): Promise<void> => {
    const { email, refreshToken, onRegistered } = this.props;
    const { additionalEmails } = this.state;

    this.setState({ submitting: true, error: null });

    try {
      const resp = await fetch("/api/register", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          email,
          refresh_token: refreshToken,
          additional_emails: additionalEmails,
        }),
      });

      if (!resp.ok) {
        const data = await resp.json();
        throw new Error(data.error || `Registration failed (${resp.status})`);
      }

      const data = await resp.json();
      localStorage.setItem("halo_watcher_registered", data.email);
      onRegistered(additionalEmails);
    } catch (err: any) {
      this.setState({ error: err.message, submitting: false });
    }
  };

  render(): React.ReactNode {
    const { email, onSkip } = this.props;
    const { additionalEmails, newEmail, submitting, error } = this.state;

    return (
      <div style={{ padding: "12px", fontFamily: "Segoe UI, sans-serif", fontSize: "13px" }}>
        <div style={{ fontSize: "14px", fontWeight: 700, marginBottom: "12px" }}>
          🔗 Enable Email Watching
        </div>

        <div style={{ color: "#666", marginBottom: "12px", fontSize: "12px" }}>
          Your email conversations will be automatically journaled to Halo tickets.
          The watcher syncs new replies in the background — no manual work needed.
        </div>

        {error && (
          <div style={{ color: "#a80000", background: "#fde7e9", padding: "8px", borderRadius: "4px", marginBottom: "12px", fontSize: "12px" }}>
            {error}
          </div>
        )}

        <div style={{ marginBottom: "8px" }}>
          <label style={{ fontSize: "12px", fontWeight: 600 }}>Your email:</label>
          <input style={inputStyle} value={email} disabled />
        </div>

        <div style={{ marginBottom: "12px" }}>
          <label style={{ fontSize: "12px", fontWeight: 600 }}>
            Additional mailboxes (optional):
          </label>
          <div style={{ display: "flex", gap: "4px" }}>
            <input
              style={inputStyle}
              placeholder="shared@yourcompany.com"
              value={newEmail}
              onChange={(e) => this.setState({ newEmail: e.target.value })}
              onKeyDown={(e) => e.key === "Enter" && this.handleAddEmail()}
            />
            <button
              style={{ ...buttonStyle, marginRight: 0, flexShrink: 0 }}
              onClick={this.handleAddEmail}
            >
              + Add
            </button>
          </div>
          {additionalEmails.map((e) => (
            <div
              key={e}
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "4px 8px",
                background: "#f0f0f0",
                borderRadius: "4px",
                marginBottom: "4px",
                fontSize: "12px",
              }}
            >
              <span>{e}</span>
              <button
                style={{ background: "none", border: "none", cursor: "pointer", color: "#a80000", fontSize: "14px" }}
                onClick={() => this.handleRemoveEmail(e)}
              >
                ✕
              </button>
            </div>
          ))}
        </div>

        <div style={{ marginTop: "12px" }}>
          <button style={buttonStyle} onClick={this.handleSubmit} disabled={submitting}>
            {submitting ? "Saving…" : "Save & Start Watching"}
          </button>
          <button style={secondaryButtonStyle} onClick={onSkip}>
            Skip
          </button>
        </div>
      </div>
    );
  }
}