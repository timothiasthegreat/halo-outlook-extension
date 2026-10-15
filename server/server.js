"use strict";

/**
 * Express server for the Halo Outlook Extension.
 *
 * Serves the add-in's static files and exposes API endpoints for
 * tenant config, mailbox registration, and health checks.
 *
 * All state.db writes are proxied to the Python watcher's FastAPI
 * server (port 8888) so only a single aiosqlite connection touches
 * the database — avoiding sql.js ↔ aiosqlite WAL corruption.
 *
 * Usage:
 *   node server.js [--port PORT] [--config PATH]
 *   Defaults: port 3000, config ../config.yaml
 */

const crypto = require("crypto");
const express = require("express");
const path = require("path");
const fs = require("fs");
const yaml = require("js-yaml");

// ── Config loading ───────────────────────────────────────────────────────

function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { port: 3000, configPath: path.resolve(__dirname, "..", "config.yaml") };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--port" && i + 1 < args.length) {
      opts.port = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === "--config" && i + 1 < args.length) {
      opts.configPath = path.resolve(args[i + 1]);
      i++;
    }
  }
  return opts;
}

function loadTenantConfig(configPath) {
  const raw = fs.readFileSync(configPath, "utf8");
  const config = yaml.load(raw);
  return {
    haloUrl: config.halo.instance_url,
    haloClientId: config.halo.client_id,
    actions: {
      emailReceived: config.halo.actions?.email_received ?? 0,
      emailSent: config.halo.actions?.email_sent ?? 16,
    },
    customFieldConvId: config.halo.custom_field_conv_id ?? 285,
    defaultTicketTypeId: config.halo.default_ticket_type_id ?? 1,
    exclusions: {
      ticketTypeIdsCreate: config.halo.exclusions?.ticket_type_ids_create ?? [],
      ticketTypeIdsSearch: config.halo.exclusions?.ticket_type_ids_search ?? [],
      statusIdsSearch: config.halo.exclusions?.status_ids_search ?? [],
    },
  };
}

// ── Watcher proxy helper ─────────────────────────────────────────────────

const WATCHER_URL = "http://127.0.0.1:8888";

/**
 * Forward a request to the Python watcher's FastAPI endpoint.
 * @param {string} route - e.g. "/track-conversation"
 * @param {object|undefined} body - JSON payload (omit for GET)
 * @param {string} label - log label
 * @param {string} [method] - HTTP method, default POST
 * @returns {Promise<{ok: boolean, status: number, body: any}>}
 */
async function forwardToWatcher(route, body, label, method = "POST") {
  const url = WATCHER_URL + route;
  const start = Date.now();
  try {
    const opts = { method };
    if (body !== undefined) {
      opts.headers = { "Content-Type": "application/json" };
      opts.body = JSON.stringify(body);
    }
    const resp = await fetch(url, opts);
    const elapsed = Date.now() - start;
    const data = await resp.json().catch(() => ({}));
    console.log(`[watcher] ${method} ${route} → ${resp.status} (${elapsed}ms)`);
    return { ok: resp.ok, status: resp.status, body: data };
  } catch (err) {
    const elapsed = Date.now() - start;
    console.error(`[watcher] ${method} ${route} → ERROR (${elapsed}ms):`, err.message);
    return { ok: false, status: 502, body: { error: "Watcher service unreachable", detail: err.message } };
  }
}

// ── Caller authentication ────────────────────────────────────────────────

/**
 * Identify the caller from their Halo access token (Authorization: Bearer).
 * Halo's /api/agent/me tells us which agent the token belongs to; the
 * result is cached per token (by SHA-256) for 5 minutes.
 *
 * Sets req.agentEmail (lowercased) or responds 401/403/502.
 */
const agentCache = new Map();
const AGENT_CACHE_MS = 5 * 60 * 1000;

async function requireAgent(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Authorization header required" });
  }
  const token = authHeader.slice(7);
  const cacheKey = crypto.createHash("sha256").update(token).digest("hex");
  const now = Date.now();
  const cached = agentCache.get(cacheKey);
  if (cached && now - cached.at < AGENT_CACHE_MS) {
    req.agentEmail = cached.email;
    return next();
  }

  try {
    const cfg = loadTenantConfig(parseArgs().configPath);
    const resp = await fetch(`${cfg.haloUrl}/api/agent/me`, {
      headers: { Accept: "application/json", Authorization: `Bearer ${token}` },
    });
    if (resp.status === 401 || resp.status === 403) {
      return res.status(401).json({ error: "Halo rejected the access token" });
    }
    if (!resp.ok) {
      return res.status(502).json({ error: `Halo agent lookup failed (${resp.status})` });
    }
    const agent = await resp.json();
    const email = String(agent.email ?? "").trim().toLowerCase();
    if (!email) {
      return res.status(403).json({ error: "Your Halo agent record has no email address" });
    }

    if (agentCache.size > 500) {
      for (const [key, entry] of agentCache) {
        if (now - entry.at >= AGENT_CACHE_MS) agentCache.delete(key);
      }
    }
    agentCache.set(cacheKey, { at: now, email });
    req.agentEmail = email;
    next();
  } catch (err) {
    console.error("[auth] Halo agent lookup error:", err.message);
    res.status(502).json({ error: "Halo API unreachable", detail: err.message });
  }
}

// ── Validation helpers ───────────────────────────────────────────────────

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_TOKEN_LENGTH = 8192;

function validateEmail(value, field) {
  if (!value || typeof value !== "string" || !EMAIL_RE.test(value.trim())) {
    return `${field} must be a valid email address`;
  }
  return null;
}

// ── App ──────────────────────────────────────────────────────────────────

// Proxied POST /Tickets and /Actions carry whole HTML email bodies (inline
// images included), well past express.json()'s 100kb default.
const JSON_BODY_LIMIT = "25mb";

const app = express();
app.use(express.json({ limit: JSON_BODY_LIMIT }));

// Serve add-in static files
const staticDir = path.resolve(__dirname, "..", "add-in", "dist");
app.use(express.static(staticDir));

// ── Routes ───────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

app.get("/api/config", (_req, res) => {
  try {
    const config = loadTenantConfig(parseArgs().configPath);
    res.json(config);
  } catch (err) {
    res.status(500).json({ error: "Failed to load configuration", detail: err.message });
  }
});

/**
 * Fetch ticket types from Halo's API, filtered to those usable for tickets.
 * Uses the agent's own OAuth token so results respect their permissions.
 *
 * Cached per-agent (by token hash) for 5 minutes.
 */
const ticketTypesCaches = new Map();
const TICKET_TYPES_CACHE_MS = 5 * 60 * 1000;

app.get("/api/ticket-types", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Authorization header required" });
  }
  const token = authHeader.slice(7);

  const cacheKey = token.slice(0, 20);
  const cached = ticketTypesCaches.get(cacheKey);
  const now = Date.now();
  if (cached && (now - cached.at) < TICKET_TYPES_CACHE_MS) {
    return res.json(cached.types);
  }

  const cfg = loadTenantConfig(parseArgs().configPath);
  try {
    const url = `${cfg.haloUrl}/api/TicketType?$filter=use eq 'tickets' and visible eq true&$select=id,name`;
    const resp = await fetch(url, {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
    });
    if (!resp.ok) {
      const text = await resp.text();
      console.error(`[ticket-types] Halo returned ${resp.status}: ${text.slice(0, 200)}`);
      return res.status(502).json({ error: "Failed to fetch ticket types from Halo" });
    }
    const data = await resp.json();
    const types = Array.isArray(data) ? data : (data.ticket_types || data.records || []);
    const result = types.map((t) => ({ id: t.id, name: t.name })).sort((a, b) => a.name.localeCompare(b.name));
    ticketTypesCaches.set(cacheKey, { at: now, types: result });
    res.json(result);
  } catch (err) {
    console.error("[ticket-types] Error fetching from Halo:", err.message);
    res.status(502).json({ error: "Ticket types unavailable", detail: err.message });
  }
});

/**
 * Proxy Halo API requests through the Express server.
 *
 * The add-in browser context cannot directly call Halo's API unless
 * service.firesideit.ca is in the manifest's <AppDomains>. By routing
 * through this proxy, the add-in only needs access to its own origin.
 *
 * Usage: POST /api/proxy/Tickets  (same path and method as Halo API)
 */
app.all("/api/proxy/*", async (req, res) => {
  const start = Date.now();
  const proxiedPath = req.params[0] ?? "";
  const normalizedPath = proxiedPath.startsWith("/") ? proxiedPath : `/${proxiedPath}`;
  const haloPath = `/api${normalizedPath}`;
  const cfg = loadTenantConfig(parseArgs().configPath);

  const queryIndex = (req.url ?? "").indexOf("?");
  const queryString = queryIndex >= 0 ? req.url.slice(queryIndex) : "";
  const url = `${cfg.haloUrl}${haloPath}${queryString}`;
  const method = req.method;

  const headers = {
    Accept: "application/json",
  };
  if (req.headers.authorization) {
    headers["Authorization"] = req.headers.authorization;
  }
  if (req.headers["content-type"]) {
    headers["Content-Type"] = req.headers["content-type"];
  }

  console.log(`[proxy] ${method} ${haloPath}${queryString}`);

  try {
    const fetchOpts = { method, headers };
    if (method !== "GET" && method !== "HEAD" && req.body) {
      fetchOpts["body"] = JSON.stringify(req.body);
      const bodyPreview = JSON.stringify(req.body);
      console.log(`[proxy]   → body: ${bodyPreview.slice(0, 500)}`);
    }
    const upstream = await fetch(url, fetchOpts);
    const elapsed = Date.now() - start;

    const body = await upstream.text();
    console.log(`[proxy] ${method} ${haloPath} → ${upstream.status} (${elapsed}ms)`);
    if (!upstream.ok && body.length < 1000) {
      console.log(`[proxy]   body: ${body}`);
    }

    res.status(upstream.status);
    if (upstream.headers.get("content-type")) {
      res.set("Content-Type", upstream.headers.get("content-type"));
    }
    res.send(body);
  } catch (err) {
    const elapsed = Date.now() - start;
    console.error(`[proxy] ${method} ${haloPath} → ERROR (${elapsed}ms):`, err.message);
    res.status(502).json({ error: "Halo API unreachable", detail: err.message });
  }
});

/**
 * Track a conversation for the watcher — proxied to Python watcher.
 *
 * The add-in calls this after linking a ticket to a conversation.
 * The request is forwarded to the Python watcher's FastAPI endpoint
 * which writes to state.db through its single aiosqlite connection.
 *
 * The conversation is always watched from the caller's own mailbox
 * (the agent identified by their Halo token), never one named in the body.
 */
app.post("/api/conversations", requireAgent, async (req, res) => {
  const { conversationId, ticketId } = req.body;

  if (!conversationId || typeof conversationId !== "string" || conversationId.trim() === "") {
    return res.status(400).json({ error: "conversationId is required and must be a non-empty string" });
  }
  if (!ticketId || typeof ticketId !== "number" || ticketId <= 0) {
    return res.status(400).json({ error: "ticketId is required and must be a positive number" });
  }

  const { ok, status, body } = await forwardToWatcher("/track-conversation", {
    conversationId: conversationId.trim(),
    ticketId,
    watchedBy: req.agentEmail,
  }, "track-conversation");

  res.status(status).json(body);
});

/**
 * Look up whether a conversation is tracked for the caller's mailbox —
 * proxied to the Python watcher's state.db.
 *
 * GET /api/conversations?conversationId=...
 * 200 { conversationId, ticketId, lastSyncAt, isStale } or 404 if untracked.
 * conversationId goes in the query string because it can contain "/".
 */
app.get("/api/conversations", requireAgent, async (req, res) => {
  const conversationId = req.query.conversationId;
  if (!conversationId || typeof conversationId !== "string" || conversationId.trim() === "") {
    return res.status(400).json({ error: "conversationId query parameter is required" });
  }
  const route =
    `/conversation?conversation_id=${encodeURIComponent(conversationId.trim())}` +
    `&watched_by=${encodeURIComponent(req.agentEmail)}`;
  const { status, body } = await forwardToWatcher(route, undefined, "conversation", "GET");
  res.status(status).json(body);
});

/**
 * Stop watching a conversation (the add-in's Unlink) — proxied to the
 * Python watcher, which removes it from state.db. Only the mailbox that
 * watches the conversation can untrack it.
 *
 * DELETE /api/conversations?conversationId=...
 */
app.delete("/api/conversations", requireAgent, async (req, res) => {
  const conversationId = req.query.conversationId;
  if (!conversationId || typeof conversationId !== "string" || conversationId.trim() === "") {
    return res.status(400).json({ error: "conversationId query parameter is required" });
  }
  const { status, body } = await forwardToWatcher("/untrack-conversation", {
    conversationId: conversationId.trim(),
    watchedBy: req.agentEmail,
  }, "untrack-conversation");
  res.status(status).json(body);
});

/**
 * Register the caller's mailbox for the watcher — proxied to Python watcher.
 *
 * The add-in calls this with a refresh token from a Halo authorization
 * dedicated to the watcher (not the add-in's own token). The caller must
 * send their Halo access token, and may only register their own email;
 * the watcher also checks that the refresh token belongs to that agent.
 */
app.post("/api/register", requireAgent, async (req, res) => {
  const { email, refresh_token, additional_emails } = req.body;

  const emailErr = validateEmail(email, "email");
  if (emailErr) {
    return res.status(400).json({ error: emailErr });
  }
  if (email.trim().toLowerCase() !== req.agentEmail) {
    return res.status(403).json({
      error: `You can only register your own mailbox (${req.agentEmail})`,
    });
  }
  if (Array.isArray(additional_emails) && additional_emails.length > 0) {
    return res.status(400).json({
      error: "additional_emails is not supported — each mailbox must register with its own Halo sign-in",
    });
  }
  if (!refresh_token || typeof refresh_token !== "string" || refresh_token.trim() === "") {
    return res.status(400).json({ error: "refresh_token is required and must be a non-empty string" });
  }
  if (refresh_token.length > MAX_TOKEN_LENGTH) {
    return res.status(400).json({ error: `refresh_token exceeds maximum length of ${MAX_TOKEN_LENGTH} characters` });
  }
  if (!/^[A-Za-z0-9\-_=]{40,}$/.test(refresh_token.trim())) {
    return res.status(400).json({ error: "refresh_token does not appear to be a valid token" });
  }

  const { ok, status, body } = await forwardToWatcher("/register-mailbox", {
    email: req.agentEmail,
    refresh_token: refresh_token.trim(),
  }, "register-mailbox");

  res.status(status).json(body);
});

/**
 * Whether the caller's mailbox is registered with a usable watcher token.
 * The add-in uses this to decide whether to offer "Enable background sync".
 */
app.get("/api/register/status", requireAgent, async (req, res) => {
  const route = `/mailbox-status?email=${encodeURIComponent(req.agentEmail)}`;
  const { status, body } = await forwardToWatcher(route, undefined, "mailbox-status", "GET");
  res.status(status).json(body);
});

// ── Export for testing ───────────────────────────────────────────────────

module.exports = { app };

// ── Start (when run directly, not required in tests) ─────────────────────

if (require.main === module) {
  const opts = parseArgs();
  app.listen(opts.port, () => {
    console.log(`Halo Outlook Server listening on http://0.0.0.0:${opts.port}`);
    console.log(`Serving add-in from: ${staticDir}`);
    console.log(`Config path: ${opts.configPath}`);
    console.log(`Watcher API: ${WATCHER_URL}`);
  });
}