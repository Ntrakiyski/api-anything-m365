import test from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import { get } from "node:http";
import { createApiServer, loadWebAssets } from "./http-api.mjs";

const token = "c".repeat(64);
async function setup(t) {
  const calls = [];
  const connection = {
    isConnecting: false,
    snapshot: () => ({ status: "connected", message: "Ready.", canRead: true, busy: false }),
    start() { calls.push("connect"); return { status: "refreshing", canRead: false, busy: true }; },
    cancel() { calls.push("cancel"); return { status: "cancelling", canRead: false, busy: true }; },
  };
  const reader = { async emails(folder, count, includeBody) {
    calls.push({ folder, count, includeBody });
    return { folder, requestedCount: count, returnedCount: 0, newestFirst: true, hasMore: false, emails: [] };
  } };
  const server = createApiServer(reader, token, { connection, assets: loadWebAssets() });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  t.after(() => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }));
  const base = `http://127.0.0.1:${server.address().port}`;
  const request = (path, options) => fetch(`${base}${path}`, options);
  const home = await request("/");
  const cookie = home.headers.get("set-cookie").split(";")[0];
  const stateResponse = await request("/ui/status", { headers: { cookie } });
  const state = await stateResponse.json();
  const headers = { cookie, origin: base, "content-type": "application/json", "x-csrf-token": state.csrfToken };
  return { base, request, cookie, headers, calls, state, home, connection };
}

test("the HTML app is served locally with HttpOnly strict cookies and no embedded API token", async t => {
  const { home, request } = await setup(t);
  const html = await home.text();
  assert.match(html, /Folder Reader/);
  assert.ok(!html.includes(token));
  assert.match(home.headers.get("set-cookie"), /HttpOnly; SameSite=Strict/);
  assert.match(home.headers.get("content-security-policy"), /script-src 'self'/);
  assert.match(home.headers.get("content-security-policy"), /frame-ancestors 'none'/);
  assert.equal(home.headers.get("cache-control"), "no-store");
  for (const path of ["/app.css", "/app.js"]) {
    const response = await request(path);
    assert.equal(response.status, 200);
    assert.equal(response.headers.get("x-content-type-options"), "nosniff");
  }
});

test("a local browser session can read without exposing an API bearer token to JavaScript", async t => {
  const { request, cookie, calls } = await setup(t);
  const response = await request("/outlook-folder/count=3?folder=Inbox&includeBody=false", { headers: { cookie } });
  assert.equal(response.status, 200);
  assert.deepEqual(calls, [{ folder: "Inbox", count: 3, includeBody: false }]);
  const badBearer = await request("/outlook-folder/count=3?folder=Inbox", {
    headers: { cookie, authorization: "Bearer wrong" },
  });
  assert.equal(badBearer.status, 401);
});

test("UI status requires a browser session and never returns Microsoft credentials", async t => {
  const { request, state } = await setup(t);
  assert.equal((await request("/ui/status")).status, 401);
  assert.equal((await request("/ui/status", { headers: { authorization: `Bearer ${token}` } })).status, 401);
  assert.match(state.csrfToken, /^[a-f0-9]{64}$/);
  assert.deepEqual(Object.keys(state).sort(), ["busy", "canRead", "csrfToken", "message", "status"]);
});

test("connect and cancel require a same-origin action with valid cookie and CSRF token", async t => {
  const { request, headers, calls } = await setup(t);
  assert.equal((await request("/ui/login", { method: "POST", headers, body: "{}" })).status, 202);
  assert.equal((await request("/ui/login/cancel", { method: "POST", headers, body: "{}" })).status, 202);
  assert.deepEqual(calls, ["connect", "cancel"]);
});

test("missing or forged Origin, CSRF, or session cannot trigger Microsoft sign-in", async t => {
  const { request, headers, calls } = await setup(t);
  for (const changed of [
    { origin: "" }, { origin: "https://example.test" }, { "x-csrf-token": "" },
    { "x-csrf-token": "wrong" }, { cookie: "" }, { cookie: "outlook_ui=forged" },
  ]) {
    const response = await request("/ui/login", { method: "POST", headers: { ...headers, ...changed }, body: "{}" });
    assert.ok([401, 403].includes(response.status));
  }
  assert.deepEqual(calls, []);
});

test("cross-site browser requests cannot issue sessions or read even with a cookie", async t => {
  const { request, cookie } = await setup(t);
  for (const path of ["/", "/ui/status", "/outlook-folder/count=3?folder=Inbox"]) {
    const response = await request(path, { headers: { cookie, "sec-fetch-site": "cross-site" } });
    assert.equal(response.status, 403);
  }
});

test("a link can open the app as a top-level document without allowing cross-site email navigation", async t => {
  const { base } = await setup(t);
  const headers = { "sec-fetch-site": "cross-site", "sec-fetch-mode": "navigate", "sec-fetch-dest": "document" };
  const navigate = path => new Promise((resolve, reject) => {
    get(`${base}${path}`, { headers }, response => {
      response.resume();
      resolve(response);
    }).on("error", reject);
  });
  const home = await navigate("/");
  assert.equal(home.statusCode, 200);
  assert.match(home.headers["set-cookie"][0], /HttpOnly; SameSite=Strict/);
  assert.equal((await navigate("/outlook-folder/count=3?folder=Inbox")).statusCode, 403);
  assert.equal((await navigate("/ui/status")).statusCode, 403);
});

test("auth actions reject arbitrary controls, malformed JSON, and oversized payloads", async t => {
  const { request, headers, calls } = await setup(t);
  for (const [body, expected] of [['{"action":"SendItem"}', 400], ["not-json", 400], ["[]", 400], ["null", 400], ["x".repeat(257), 413]]) {
    assert.equal((await request("/ui/login", { method: "POST", headers, body })).status, expected);
  }
  assert.equal((await request("/ui/login", {
    method: "POST", headers: { ...headers, "content-type": "text/plain" }, body: "{}",
  })).status, 400);
  assert.deepEqual(calls, []);
});

test("email reads remain GET-only and pause while Microsoft sign-in is active", async t => {
  const { request, cookie, connection } = await setup(t);
  assert.equal((await request("/outlook-folder/count=3?folder=Inbox", {
    method: "POST", headers: { cookie }, body: "{}",
  })).status, 405);
  connection.isConnecting = true;
  const response = await request("/outlook-folder/count=3?folder=Inbox", { headers: { cookie } });
  assert.equal(response.status, 409);
  assert.equal((await response.json()).error, "AUTH_IN_PROGRESS");
});
