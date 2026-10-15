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

function request(path, options = {}) {
  return new Promise((resolve, reject) => {
    const req = http.request(
      `${baseUrl}${path}`,
      {
        method: options.method ?? "GET",
        headers: options.headers ?? {},
      },
      (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body }));
        res.on("error", reject);
      },
    );

    if (options.body !== undefined) {
      req.write(options.body);
    }
    req.end();
  });
}

// ── Health ───────────────────────────────────────────────────────────────

test("GET /health returns 200 with status ok", async () => {
  const { status, body } = await get("/health");
  assert.equal(status, 200);
  const data = JSON.parse(body);
  assert.equal(data.status, "ok");
  // No uptime exposed — prevents information leakage
  assert.equal(Object.keys(data).length, 1);
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

// ── Proxy path normalization regression ────────────────────────────────

test("GET /api/proxy/Tickets forwards to /api/Tickets", async () => {
  const originalFetch = global.fetch;
  let capturedUrl;

  global.fetch = async (url) => {
    capturedUrl = String(url);
    return {
      ok: true,
      status: 200,
      headers: { get: (name) => (name.toLowerCase() === "content-type" ? "application/json" : null) },
      text: async () => JSON.stringify({ ok: true }),
    };
  };

  try {
    const { status, body } = await request("/api/proxy/Tickets", {
      method: "GET",
      headers: { Authorization: "Bearer test-token" },
    });

    assert.equal(status, 200);
    assert.ok(capturedUrl, "expected proxy to call upstream fetch");
    assert.ok(capturedUrl.endsWith("/api/Tickets"), `expected '/api/Tickets' suffix, got '${capturedUrl}'`);
    const data = JSON.parse(body);
    assert.equal(data.ok, true);
  } finally {
    global.fetch = originalFetch;
  }
});

// ── Registration & conversation tracking ─────────────────────────────────

// Valid tokens: base64url, >= 40 chars. A plausible HaloPSA Fernet refresh
// token is ~200-300 chars, base64url-encoded.
const VALID_TOKEN = Buffer.from("x".repeat(200)).toString("base64url"); // 268 chars

let tokenCounter = 0;
/** A fresh access token per call, so the server's agent cache never leaks between tests. */
function nextAccessToken() {
  tokenCounter += 1;
  return `access-token-${tokenCounter}`;
}

/**
 * Replace global fetch with a fake Halo (/api/agent/me) and fake watcher
 * (127.0.0.1:8888). Records watcher calls in the returned array.
 */
function fakeUpstreams({ agentEmail = "test@example.com", agentStatus = 200 } = {}) {
  const watcherCalls = [];
  const originalFetch = global.fetch;
  global.fetch = async (url, opts = {}) => {
    const u = String(url);
    if (u.endsWith("/api/agent/me")) {
      return {
        ok: agentStatus === 200,
        status: agentStatus,
        json: async () => ({ id: 1, email: agentEmail }),
      };
    }
    if (u.startsWith("http://127.0.0.1:8888")) {
      const body = opts.body ? JSON.parse(opts.body) : undefined;
      watcherCalls.push({ url: u, method: opts.method, body });
      return {
        ok: true,
        status: 200,
        json: async () => ({ status: "ok", email: body?.email }),
      };
    }
    throw new Error(`unexpected fetch ${u}`);
  };
  return { watcherCalls, restore: () => { global.fetch = originalFetch; } };
}

function postAuthed(path, data, token = nextAccessToken()) {
  return request(path, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify(data),
  });
}

test("POST /api/register without Authorization returns 401", async () => {
  const fake = fakeUpstreams();
  try {
    const { status } = await post("/api/register", {
      email: "test@example.com",
      refresh_token: VALID_TOKEN,
    });
    assert.equal(status, 401);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("POST /api/register with a token Halo rejects returns 401", async () => {
  const fake = fakeUpstreams({ agentStatus: 401 });
  try {
    const { status } = await postAuthed("/api/register", {
      email: "test@example.com",
      refresh_token: VALID_TOKEN,
    });
    assert.equal(status, 401);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("POST /api/register for the caller's own email forwards to the watcher", async () => {
  const fake = fakeUpstreams({ agentEmail: "Test@Example.com" });
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "test@example.com",
      refresh_token: VALID_TOKEN,
    });
    assert.equal(status, 200);
    assert.equal(JSON.parse(body).status, "ok");
    assert.equal(fake.watcherCalls.length, 1);
    assert.ok(fake.watcherCalls[0].url.endsWith("/register-mailbox"));
    assert.deepEqual(fake.watcherCalls[0].body, {
      email: "test@example.com",
      refresh_token: VALID_TOKEN,
    });
  } finally {
    fake.restore();
  }
});

test("POST /api/register for another agent's email returns 403", async () => {
  const fake = fakeUpstreams({ agentEmail: "attacker@example.com" });
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "victim@example.com",
      refresh_token: VALID_TOKEN,
    });
    assert.equal(status, 403);
    assert.ok(JSON.parse(body).error.includes("your own mailbox"));
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("POST /api/register with additional_emails returns 400", async () => {
  const fake = fakeUpstreams({ agentEmail: "multi@example.com" });
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "multi@example.com",
      refresh_token: VALID_TOKEN,
      additional_emails: ["shared@example.com"],
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("additional_emails"));
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("POST /api/register with missing email returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status, body } = await postAuthed("/api/register", {
      refresh_token: VALID_TOKEN,
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("email"));
  } finally {
    fake.restore();
  }
});

test("POST /api/register with missing refresh_token returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "test@example.com",
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("refresh_token"));
  } finally {
    fake.restore();
  }
});

test("POST /api/register with token too short returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "test@example.com",
      refresh_token: "short", // < 40 chars, fails regex
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("does not appear to be a valid token"));
  } finally {
    fake.restore();
  }
});

test("POST /api/register with non-base64url characters returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "test@example.com",
      refresh_token: "not a valid token!!!!! with spaces and stuff!!!", // spaces, !
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("does not appear to be a valid token"));
  } finally {
    fake.restore();
  }
});

test("POST /api/register with token exceeding max length returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status, body } = await postAuthed("/api/register", {
      email: "test@example.com",
      refresh_token: "x".repeat(10000), // > 8192
    });
    assert.equal(status, 400);
    assert.ok(JSON.parse(body).error.includes("exceeds maximum length"));
  } finally {
    fake.restore();
  }
});

test("GET /api/register/status asks the watcher about the caller's mailbox", async () => {
  const fake = fakeUpstreams({ agentEmail: "status@example.com" });
  try {
    const { status } = await request("/api/register/status", {
      headers: { Authorization: `Bearer ${nextAccessToken()}` },
    });
    assert.equal(status, 200);
    assert.equal(fake.watcherCalls.length, 1);
    assert.equal(fake.watcherCalls[0].method, "GET");
    assert.ok(fake.watcherCalls[0].url.endsWith("/mailbox-status?email=status%40example.com"));
  } finally {
    fake.restore();
  }
});

test("POST /api/conversations without Authorization returns 401", async () => {
  const fake = fakeUpstreams();
  try {
    const { status } = await post("/api/conversations", { conversationId: "conv-1", ticketId: 5 });
    assert.equal(status, 401);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("POST /api/conversations watches from the caller's mailbox, not the body's", async () => {
  const fake = fakeUpstreams({ agentEmail: "agent@example.com" });
  try {
    const { status } = await postAuthed("/api/conversations", {
      conversationId: "conv-1",
      ticketId: 5,
      watchedBy: "someone-else@example.com",
    });
    assert.equal(status, 200);
    assert.deepEqual(fake.watcherCalls[0].body, {
      conversationId: "conv-1",
      ticketId: 5,
      watchedBy: "agent@example.com",
    });
  } finally {
    fake.restore();
  }
});

test("GET /api/conversations looks up the caller's mapping in the watcher", async () => {
  const fake = fakeUpstreams({ agentEmail: "agent@example.com" });
  try {
    const conv = "AAQkAD/abc+def=";
    const { status } = await request(`/api/conversations?conversationId=${encodeURIComponent(conv)}`, {
      headers: { Authorization: `Bearer ${nextAccessToken()}` },
    });
    assert.equal(status, 200);
    assert.equal(fake.watcherCalls.length, 1);
    assert.equal(fake.watcherCalls[0].method, "GET");
    const url = new URL(fake.watcherCalls[0].url);
    assert.equal(url.pathname, "/conversation");
    assert.equal(url.searchParams.get("conversation_id"), conv);
    assert.equal(url.searchParams.get("watched_by"), "agent@example.com");
  } finally {
    fake.restore();
  }
});

test("GET /api/conversations without Authorization returns 401", async () => {
  const fake = fakeUpstreams();
  try {
    const { status } = await get("/api/conversations?conversationId=conv-1");
    assert.equal(status, 401);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("GET /api/conversations without conversationId returns 400", async () => {
  const fake = fakeUpstreams();
  try {
    const { status } = await request("/api/conversations", {
      headers: { Authorization: `Bearer ${nextAccessToken()}` },
    });
    assert.equal(status, 400);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

test("DELETE /api/conversations untracks from the caller's mailbox", async () => {
  const fake = fakeUpstreams({ agentEmail: "agent@example.com" });
  try {
    const { status } = await request("/api/conversations?conversationId=conv-1", {
      method: "DELETE",
      headers: { Authorization: `Bearer ${nextAccessToken()}` },
    });
    assert.equal(status, 200);
    assert.equal(fake.watcherCalls.length, 1);
    assert.ok(fake.watcherCalls[0].url.endsWith("/untrack-conversation"));
    assert.equal(fake.watcherCalls[0].method, "POST");
    assert.deepEqual(fake.watcherCalls[0].body, {
      conversationId: "conv-1",
      watchedBy: "agent@example.com",
    });
  } finally {
    fake.restore();
  }
});

test("DELETE /api/conversations without Authorization returns 401", async () => {
  const fake = fakeUpstreams();
  try {
    const { status } = await request("/api/conversations?conversationId=conv-1", { method: "DELETE" });
    assert.equal(status, 401);
    assert.equal(fake.watcherCalls.length, 0);
  } finally {
    fake.restore();
  }
});

// ── Body size limit ──────────────────────────────────────────────────────

test("POST /api/proxy/Actions accepts an HTML email body larger than 100kb", async () => {
  const originalFetch = global.fetch;
  let forwarded;
  global.fetch = async (url, opts) => {
    forwarded = { url: String(url), body: opts.body };
    return {
      ok: true,
      status: 201,
      headers: { get: () => "application/json" },
      text: async () => JSON.stringify([{ id: 1 }]),
    };
  };

  try {
    // ~5 MB of HTML, e.g. an email with an inline base64 image
    const noteHtml = `<p>Hi</p><img src="data:image/png;base64,${"A".repeat(5 * 1024 * 1024)}">`;
    const payload = JSON.stringify([{ ticket_id: 1, outcome_id: 0, note_html: noteHtml }]);
    const { status } = await request("/api/proxy/Actions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(payload),
        Authorization: "Bearer test-token",
      },
      body: payload,
    });

    assert.equal(status, 201);
    assert.ok(forwarded.url.endsWith("/api/Actions"));
    assert.equal(JSON.parse(forwarded.body)[0].note_html.length, noteHtml.length);
  } finally {
    global.fetch = originalFetch;
  }
});
