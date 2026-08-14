"use strict";

/**
 * Express server for the Halo Outlook Extension.
 *
 * Serves the add-in's static files and exposes API endpoints for
 * tenant config, mailbox registration, and health checks.
 *
 * Usage:
 *   node server.js [--port PORT] [--config PATH]
 *   Defaults: port 3000, config ../config.yaml
 */

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

function loadConfig(configPath) {
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

function validateEmail(value, field) {
  if (!value || typeof value !== "string" || !EMAIL_RE.test(value.trim())) {
    return `${field} must be a valid email address`;
  }
  return null;
}

// ── Registration store (Phase 2 moves this to state.db) ──────────────────
// For now, in-memory store so registration endpoint is functional for testing.

const registrationStore = new Map();

// ── App ──────────────────────────────────────────────────────────────────

const app = express();
app.use(express.json());

// Serve add-in static files
const staticDir = path.resolve(__dirname, "..", "add-in", "dist");
app.use(express.static(staticDir));

// ── Routes ───────────────────────────────────────────────────────────────

const startTime = Date.now();

app.get("/health", (_req, res) => {
  res.json({ status: "ok", uptime: Math.floor((Date.now() - startTime) / 1000) });
});

app.get("/api/config", (_req, res) => {
  try {
    const config = loadConfig(parseArgs().configPath);
    res.json(config);
  } catch (err) {
    res.status(500).json({ error: "Failed to load configuration", detail: err.message });
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

  // Store registration (Phase 2: encrypt token and write to state.db)
  registrationStore.set(normalizedEmail, {
    email: normalizedEmail,
    refresh_token,
    additional_emails: additional.map((e) => e.trim().toLowerCase()),
    registered_at: new Date().toISOString(),
  });

  res.json({ status: "ok", email: normalizedEmail, watching });
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
  });
}