"use strict";

/**
 * Express server for the Halo Outlook Extension.
 *
 * Serves the add-in's static files and exposes API endpoints for
 * tenant config, mailbox registration, and health checks.
 *
 * Stores registrations in the watcher's state.db (shared SQLite volume)
 * with Fernet-encrypted refresh tokens.
 *
 * Usage:
 *   node server.js [--port PORT] [--config PATH]
 *   Defaults: port 3000, config ../config.yaml
 */

const express = require("express");
const path = require("path");
const fs = require("fs");
const yaml = require("js-yaml");
const initSqlJs = require("sql.js");

const { createFernet } = require("./fernet");

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
  };
}

// ── Validation helpers ───────────────────────────────────────────────────

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_TOKEN_LENGTH = 8192; // OAuth refresh tokens are typically < 2 KB

function validateEmail(value, field) {
  if (!value || typeof value !== "string" || !EMAIL_RE.test(value.trim())) {
    return `${field} must be a valid email address`;
  }
  return null;
}

// ── Database initialization ──────────────────────────────────────────────

/** @type {import("sql.js").Database | null} */
let db = null;
let fernet = null;

async function initDb() {
  const SQL = await initSqlJs();
  // Detect state.db path: if state_db_path is set in config, use it;
  // otherwise default to ../state.db (watcher default).
  const opts = parseArgs();
  let dbPath;
  try {
    const config = yaml.load(fs.readFileSync(opts.configPath, "utf8"));
    dbPath = config.watcher?.state_db_path || path.resolve(__dirname, "..", "state.db");
  } catch {
    dbPath = path.resolve(__dirname, "..", "state.db");
  }

  // Open or create the database
  if (fs.existsSync(dbPath)) {
    const buffer = fs.readFileSync(dbPath);
    db = new SQL.Database(buffer);
  } else {
    db = new SQL.Database();
  }
  // Ensure the table exists
  db.run(`
    CREATE TABLE IF NOT EXISTS watched_mailboxes (
      email TEXT PRIMARY KEY,
      refresh_token_enc TEXT NOT NULL,
      token_status TEXT NOT NULL DEFAULT 'active',
      registered_at TEXT NOT NULL,
      last_token_refresh_at TEXT
    )
  `);
  saveDb();
  console.log(`SQLite state.db opened at ${dbPath}`);
}

function saveDb() {
  if (!db) return;
  const data = db.export();
  const opts = parseArgs();
  let dbPath;
  try {
    const config = yaml.load(fs.readFileSync(opts.configPath, "utf8"));
    dbPath = config.watcher?.state_db_path || path.resolve(__dirname, "..", "state.db");
  } catch {
    dbPath = path.resolve(__dirname, "..", "state.db");
  }
  fs.writeFileSync(dbPath, Buffer.from(data));
}

function getFernet() {
  if (!fernet) {
    const key = process.env.FERNET_KEY;
    if (!key) {
      console.warn("FERNET_KEY not set — tokens will NOT be encrypted at rest!");
      // Fall back to a no-op pass-through (base64 encode/decode only)
      // This lets the server start for testing but logs a clear warning.
      fernet = {
        encrypt: (s) => Buffer.from(s).toString("base64"),
        decrypt: (s) => Buffer.from(s, "base64").toString("utf8"),
      };
    } else {
      fernet = createFernet(key);
    }
  }
  return fernet;
}

// ── App ──────────────────────────────────────────────────────────────────

const app = express();
app.use(express.json());

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
const ticketTypesCaches = new Map(); // key: token-hash, value: { at: timestamp, types: [...] }
const TICKET_TYPES_CACHE_MS = 5 * 60 * 1000; // 5 minutes

app.get("/api/ticket-types", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Authorization header required" });
  }
  const token = authHeader.slice(7);

  // Per-agent cache by token prefix (first 20 chars — unique enough, not the full token)
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
 * All requests are logged with timing for diagnostics.
 *
 * Usage: POST /api/proxy/Tickets  (same path and method as Halo API)
 * The add-in sends its bearer token in the Authorization header, which
 * this proxy forwards unchanged to Halo.
 */
app.all("/api/proxy/*", async (req, res) => {
  const start = Date.now();
  const proxiedPath = req.params[0] ?? "";
  const normalizedPath = proxiedPath.startsWith("/") ? proxiedPath : `/${proxiedPath}`;
  const haloPath = `/api${normalizedPath}`; // /api/proxy/Tickets -> /api/Tickets
  const cfg = loadTenantConfig(parseArgs().configPath);
  const url = `${cfg.haloUrl}${haloPath}`;
  const method = req.method;

  // Forward the request to Halo
  const headers = {
    Accept: "application/json",
  };
  if (req.headers.authorization) {
    headers["Authorization"] = req.headers.authorization;
  }
  if (req.headers["content-type"]) {
    headers["Content-Type"] = req.headers["content-type"];
  }

  console.log(`[proxy] ${method} ${haloPath}`);

  try {
    const fetchOpts = { method, headers };
    if (method !== "GET" && method !== "HEAD" && req.body) {
      fetchOpts["body"] = JSON.stringify(req.body);
    }
    const upstream = await fetch(url, fetchOpts);
    const elapsed = Date.now() - start;

    // Read the response body once for forwarding and once for logging
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

app.post("/api/register", (req, res) => {
  const { email, refresh_token, additional_emails } = req.body;

  // Validate required fields
  const emailErr = validateEmail(email, "email");
  if (emailErr) {
    return res.status(400).json({ error: emailErr });
  }
  if (!refresh_token || typeof refresh_token !== "string" || refresh_token.trim() === "") {
    return res.status(400).json({ error: "refresh_token is required and must be a non-empty string" });
  }
  if (refresh_token.length > MAX_TOKEN_LENGTH) {
    return res.status(400).json({ error: `refresh_token exceeds maximum length of ${MAX_TOKEN_LENGTH} characters` });
  }
  // Basic format check: HaloPSA refresh tokens are base64-encoded Fernet tokens
  // (~200-300 chars, no whitespace, base64url alphabet)
  if (!/^[A-Za-z0-9\-_=]{40,}$/.test(refresh_token.trim())) {
    return res.status(400).json({ error: "refresh_token does not appear to be a valid token" });
  }

  const normalizedEmail = email.trim().toLowerCase();

  // Validate additional emails if provided
  const additional = Array.isArray(additional_emails) ? additional_emails : [];
  const watching = [normalizedEmail];
  for (const extra of additional) {
    const err = validateEmail(extra, "additional_emails");
    if (err) {
      return res.status(400).json({ error: err });
    }
    watching.push(extra.trim().toLowerCase());
  }

  // Encrypt token and write to state.db
  try {
    const f = getFernet();
    const encToken = f.encrypt(refresh_token);
    const now = new Date().toISOString();

    if (!db) {
      // Fallback: in-memory store if db not initialized (tests without state.db)
      return res.json({
        status: "ok",
        email: normalizedEmail,
        watching,
        _db_initialized: false,
      });
    }

    // Upsert: register each mailbox
    const stmt = db.prepare(
      `INSERT INTO watched_mailboxes
         (email, refresh_token_enc, token_status, registered_at, last_token_refresh_at)
       VALUES (?, ?, 'active', ?, ?)
       ON CONFLICT(email) DO UPDATE SET
         refresh_token_enc = excluded.refresh_token_enc,
         token_status = 'active',
         last_token_refresh_at = excluded.last_token_refresh_at`
    );

    for (const mailbox of watching) {
      stmt.run([mailbox, encToken, now, now]);
    }
    stmt.free();
    saveDb();

    res.json({ status: "ok", email: normalizedEmail, watching });
  } catch (err) {
    console.error("Registration database error:", err);
    res.status(500).json({ error: "Failed to store registration", detail: err.message });
  }
});

// ── Export for testing ───────────────────────────────────────────────────

module.exports = { app };

// ── Start (when run directly, not required in tests) ─────────────────────

if (require.main === module) {
  const opts = parseArgs();
  initDb().then(() => {
    app.listen(opts.port, () => {
      console.log(`Halo Outlook Server listening on http://0.0.0.0:${opts.port}`);
      console.log(`Serving add-in from: ${staticDir}`);
      console.log(`Config path: ${opts.configPath}`);
    });
  }).catch((err) => {
    console.error("Failed to initialize database:", err);
    process.exit(1);
  });
}