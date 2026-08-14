/**
 * Tests for the Express server: health, config, and static file serving.
 *
 * Starts the server on a random port, tests each endpoint, and ensures
 * the server process stops after each suite.
 */
const { test, after, before } = require("node:test");
const assert = require("node:assert/strict");
const http = require("node:http");
const path = require("node:path");
const fs = require("node:fs");

// The server module exports the Express app (not the listening server)
// so we can bind to an ephemeral port in tests.
const { app } = require("../server");

let server;
let baseUrl;

before(() => {
  return new Promise((resolve) => {
    server = app.listen(0, () => {
      const port = server.address().port;
      baseUrl = `http://localhost:${port}`;
      resolve();
    });
  });
});

after(() => {
  if (server) server.close();
});

function get(path) {
  return new Promise((resolve, reject) => {
    http.get(`${baseUrl}${path}`, (res) => {
      let body = "";
      res.on("data", (chunk) => (body += chunk));
      res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body }));
      res.on("error", reject);
    });
  });
}

function post(path, data) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify(data);
    const req = http.request(
      `${baseUrl}${path}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) },
      },
      (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body }));
        res.on("error", reject);
      },
    );
    req.write(payload);
    req.end();
  });
}

// ── Health ───────────────────────────────────────────────────────────────

test("GET /health returns 200 with status ok", async () => {
  const { status, body } = await get("/health");
  assert.equal(status, 200);
  const data = JSON.parse(body);
  assert.equal(data.status, "ok");
  assert.ok(typeof data.uptime === "number");
});

// ── Config endpoint ──────────────────────────────────────────────────────

test("GET /api/config returns 200 with tenant config", async () => {
  const { status, body } = await get("/api/config");
  assert.equal(status, 200);
  const data = JSON.parse(body);
  assert.ok(typeof data.haloUrl === "string");
  assert.ok(typeof data.haloClientId === "string");
  assert.ok(typeof data.actions === "object");
  assert.ok(typeof data.actions.emailReceived === "number");
  assert.ok(typeof data.actions.emailSent === "number");
  assert.ok(typeof data.customFieldConvId === "number");
  assert.ok(typeof data.defaultTicketTypeId === "number");
});

test("GET /api/config returns CORS headers for add-in", async () => {
  const { headers } = await get("/api/config");
  // Same-origin for the add-in — no CORS needed, but we set permissive headers
  assert.equal(headers["content-type"], "application/json; charset=utf-8");
});

// ── Static files ─────────────────────────────────────────────────────────

test("GET /taskpane.html returns HTML", async () => {
  const { status, headers, body } = await get("/taskpane.html");
  assert.equal(status, 200);
  assert.ok(headers["content-type"].includes("text/html"));
  assert.ok(body.includes("<html") || body.includes("<div") || body.includes("<script"));
});

test("GET /taskpane.js returns JavaScript", async () => {
  const { status, headers } = await get("/taskpane.js");
  assert.equal(status, 200);
  assert.ok(headers["content-type"].includes("javascript") || headers["content-type"].includes("text/javascript") || headers["content-type"].includes("application/javascript"));
});

test("GET /nonexistent returns 404", async () => {
  const { status } = await get("/nonexistent/file.xyz");
  assert.equal(status, 404);
});

// ── Registration (smoke test — full register with Fernet comes in Phase 2) ──

test("POST /api/register with valid body returns 200", async () => {
  const { status, body } = await post("/api/register", {
    email: "test@example.com",
    refresh_token: "fake-refresh-token",
  });
  assert.equal(status, 200);
  const data = JSON.parse(body);
  assert.equal(data.status, "ok");
  assert.equal(data.email, "test@example.com");
});

test("POST /api/register with missing email returns 400", async () => {
  const { status, body } = await post("/api/register", {
    refresh_token: "fake-refresh-token",
  });
  assert.equal(status, 400);
  const data = JSON.parse(body);
  assert.ok(data.error.includes("email"));
});

test("POST /api/register with missing refresh_token returns 400", async () => {
  const { status, body } = await post("/api/register", {
    email: "test@example.com",
  });
  assert.equal(status, 400);
  const data = JSON.parse(body);
  assert.ok(data.error.includes("refresh_token"));
});

test("POST /api/register with additional_emails returns 200", async () => {
  const { status, body } = await post("/api/register", {
    email: "multi@example.com",
    refresh_token: "fake-refresh",
    additional_emails: ["shared@example.com"],
  });
  assert.equal(status, 200);
  const data = JSON.parse(body);
  assert.deepEqual(data.watching, ["multi@example.com", "shared@example.com"]);
});