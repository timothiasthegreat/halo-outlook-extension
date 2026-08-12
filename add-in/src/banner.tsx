/**
 * Ticket info banner component — shows ticket reference, status,
 * assigned agent, and a link to open in Halo. Always visible when
 * the current email is part of a tracked conversation.
 */

import * as React from "react";
import config from "./config";
import type { HaloTicket } from "./halo";

interface BannerProps {
  ticket: HaloTicket | null;
  loading: boolean;
  lastSyncAt: string | null;
}

const bannerStyle: React.CSSProperties = {
  background: "#f0f6ff",
  border: "1px solid #c7e0f4",
  borderRadius: "4px",
  padding: "12px",
  marginBottom: "12px",
  fontFamily: "Segoe UI, sans-serif",
  fontSize: "13px",
};

const linkStyle: React.CSSProperties = {
  color: "#0078d4",
  textDecoration: "none",
  fontWeight: 600,
  cursor: "pointer",
};

const rowStyle: React.CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: "8px",
  flexWrap: "wrap" as const,
};

const badgeStyle: React.CSSProperties = {
  display: "inline-block",
  padding: "2px 8px",
  borderRadius: "10px",
  fontSize: "11px",
  fontWeight: 600,
};

function statusBadge(statusName: string | undefined): React.CSSProperties {
  const lower = (statusName ?? "").toLowerCase();
  let bg = "#e0e0e0";
  let color = "#333";
  if (lower.includes("open") || lower.includes("new")) {
    bg = "#e6f7e6";
    color = "#0a6e0a";
  } else if (
    lower.includes("waiting") ||
    lower.includes("pending") ||
    lower.includes("scheduled")
  ) {
    bg = "#fff4ce";
    color = "#8a6d3b";
  } else if (lower.includes("closed") || lower.includes("resolved")) {
    bg = "#e0e0e0";
    color = "#666";
  }
  return { ...badgeStyle, background: bg, color };
}

const Banner: React.FC<BannerProps> = ({ ticket, loading, lastSyncAt }) => {
  // ── loading state ──
  if (loading) {
    return (
      <div style={bannerStyle}>
        <div style={{ color: "#666" }}>Loading ticket information…</div>
      </div>
    );
  }

  // ── untracked state ──
  if (!ticket) {
    return (
      <div style={bannerStyle}>
        <div style={{ fontSize: "14px", fontWeight: 600, marginBottom: "4px" }}>
          🔗 Halo Ticket Bridge
        </div>
        <div style={{ color: "#666" }}>
          This conversation is not linked to a ticket.
        </div>
      </div>
    );
  }

  // ── tracked state ──
  const ticketUrl = `${config.haloUrl.replace(/\/$/, "")}/tickets/${ticket.id}`;
  const ref = ticket.ticketnumber
    ? `TKT-${ticket.ticketnumber}`
    : `#${ticket.id}`;

  return (
    <div style={bannerStyle}>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
        }}
      >
        <div style={{ flex: 1 }}>
          <div style={rowStyle}>
            <span style={{ fontSize: "14px", fontWeight: 700 }}>📋 {ref}</span>
            <span style={statusBadge(ticket.status_name)}>
              {ticket.status_name ?? `Status ${ticket.status_id}`}
            </span>
          </div>
          <div
            style={{ color: "#333", marginTop: "4px", fontSize: "13px" }}
            title={ticket.summary}
          >
            {ticket.summary.slice(0, 100)}
            {ticket.summary.length > 100 ? "…" : ""}
          </div>
          {ticket.assigned_to_name && (
            <div style={{ color: "#666", fontSize: "12px", marginTop: "2px" }}>
              Agent: {ticket.assigned_to_name}
            </div>
          )}
          {lastSyncAt && (
            <div style={{ color: "#999", fontSize: "11px", marginTop: "2px" }}>
              Last synced: {new Date(lastSyncAt).toLocaleString()}
            </div>
          )}
        </div>
        <a
          href={ticketUrl}
          target="_blank"
          rel="noopener noreferrer"
          style={linkStyle}
        >
          Open in Halo ↗
        </a>
      </div>
    </div>
  );
};

export default Banner;