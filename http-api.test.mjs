import test from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import { request as httpRequest } from "node:http";
import { createApiServer } from "./http-api.mjs";
import { ApiError } from "./folder.mjs";

const token = "a".repeat(64);
const authorization = `Bearer ${token}`;
async function setup(t, reader) {
  const calls = [];
  const server = createApiServer(reader ?? {
    async emails(folder, count, includeBody) {
      calls.push({ folder, count, includeBody });
      return { folder, requestedCount: count, returnedCount: 0, emails: [] };
    },
  }, token);
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  t.after(() => new Promise(resolve => {
    server.close(resolve);
    server.closeAllConnections();
  }));
  const base = `http://127.0.0.1:${server.address().port}`;
  const request = (path, options = {}) => fetch(`${base}${path}`, {
    ...options, headers: { authorization, ...options.headers },
  });
  return { request, calls, base };
}

test("the single endpoint accepts a folder, a path count, and both includeBody values", async t => {
  const { request, calls } = await setup(t);
  assert.equal((await request("/outlook-folder/count=3?folder=HAMMER-HAI&includeBody=false")).status, 200);
  const full = await request("/outlook-folder/count=5?folder=Inbox&includeBody=true");
  assert.equal(full.status, 200);
  assert.equal(full.headers.get("cache-control"), "no-store");
  assert.equal(full.headers.get("access-control-allow-origin"), null);
  assert.deepEqual(calls, [
    { folder: "HAMMER-HAI", count: 3, includeBody: false },
    { folder: "Inbox", count: 5, includeBody: true },
  ]);
});

test("invalid inputs, duplicate parameters, and extra controls are rejected before reading", async t => {
  const { request, calls } = await setup(t);
  for (const path of [
    "/outlook-folder/count=3",
    "/outlook-folder/count=3?folder=",
    "/outlook-folder/count=?folder=Inbox",
    "/outlook-folder/count=0?folder=Inbox",
    "/outlook-folder/count=-1?folder=Inbox",
    "/outlook-folder/count=1.5?folder=Inbox",
    "/outlook-folder/count=03?folder=Inbox",
    "/outlook-folder/count=9007199254740992?folder=Inbox",
    "/outlook-folder/count=3?folder=Inbox&folder=Other",
    "/outlook-folder/count=3?folder=Inbox&includeBody=true&includeBody=false",
    "/outlook-folder/count=3?folder=Inbox&includeBody=",
    "/outlook-folder/count=3?folder=Inbox&includeBody=yes",
    "/outlook-folder/count=3?folder=Inbox&includeBody=True",
    "/outlook-folder/count=3?folder=Inbox&count=4",
    "/outlook-folder/count=3?folder=Inbox&folderId=other",
    "/outlook-folder/count=3?folder=Inbox&itemId=other",
    "/outlook-folder/count=3?folder=Inbox&action=DeleteItem",
    "/outlook-folder/count=3?folder=Inbox&fields=body",
    "/outlook-folder/count=3?folder=Inbox&offset=0",
  ]) {
    assert.equal((await request(path)).status, 400, path);
  }
  assert.equal(calls.length, 0);
});

test("all non-GET methods are rejected without touching Outlook", async t => {
  const { request, calls } = await setup(t);
  for (const method of ["POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS"]) {
    const result = await request("/outlook-folder/count=3?folder=Inbox", { method });
    assert.equal(result.status, 405, method);
    assert.equal(result.headers.get("allow"), "GET");
  }
  assert.equal(calls.length, 0);
});

test("email requests require the exact bearer token", async t => {
  const { request, calls } = await setup(t);
  for (const value of ["", "Bearer wrong", `Basic ${token}`, `Bearer ${"b".repeat(64)}`]) {
    const result = await request("/outlook-folder/count=3?folder=Inbox", { headers: { authorization: value } });
    assert.equal(result.status, 401);
  }
  assert.equal(calls.length, 0);
});

test("browser-origin and nonlocal Host requests are refused", async t => {
  const { request, calls, base } = await setup(t);
  assert.equal((await request("/outlook-folder/count=3?folder=Inbox",
    { headers: { origin: "https://example.test" } })).status, 403);
  const status = await new Promise((resolve, reject) => {
    const raw = httpRequest(`${base}/outlook-folder/count=3?folder=Inbox`, {
      headers: { host: "example.test:3838", authorization },
    }, response => {
      response.resume();
      resolve(response.statusCode);
    });
    raw.on("error", reject);
    raw.end();
  });
  assert.equal(status, 403);
  assert.equal(calls.length, 0);
});

test("health is liveness-only and unknown endpoints do not trigger mailbox reads", async t => {
  const { request, calls } = await setup(t);
  const health = await request("/health", { headers: { authorization: "" } });
  assert.deepEqual(await health.json(), { ok: true });
  assert.equal((await request("/other?count=3")).status, 404);
  assert.equal((await request("/hammer-hai/list-emails?count=3")).status, 404);
  assert.equal((await request("/hammer-hai/read-emails?count=3")).status, 404);
  assert.equal(calls.length, 0);
});

test("reader failures retain their explicit error and do not produce success-shaped data", async t => {
  const { request } = await setup(t, {
    async emails() { throw new ApiError("FOLDER_MISMATCH", "Folder changed.", 409); },
  });
  const result = await request("/outlook-folder/count=3?folder=Inbox");
  assert.equal(result.status, 409);
  assert.deepEqual(await result.json(), { error: "FOLDER_MISMATCH", message: "Folder changed." });
});

test("overlapping email requests fail explicitly rather than queueing unbounded reads", async t => {
  let release;
  let started;
  const begun = new Promise(resolve => { started = resolve; });
  const pending = new Promise(resolve => { release = resolve; });
  const { request } = await setup(t, {
    async emails() {
      started();
      await pending;
      return { folder: "HAMMER-HAI", returnedCount: 0, emails: [] };
    },
  });
  const first = request("/outlook-folder/count=3?folder=Inbox");
  await begun;
  try {
    assert.equal((await request("/outlook-folder/count=1?folder=Inbox&includeBody=false")).status, 429);
  } finally {
    release();
  }
  assert.equal((await first).status, 200);
});

test("includeBody defaults to true and folder names are URL-decoded", async t => {
  const { request, calls } = await setup(t);
  const folder = "Projet \u00e9t\u00e9 & IA";
  const query = new URLSearchParams({ folder });
  assert.equal((await request(`/outlook-folder/count=1?${query}`)).status, 200);
  assert.deepEqual(calls, [{ folder, count: 1, includeBody: true }]);
});

test("a missing folder returns the exact requested 404 message", async t => {
  const { request } = await setup(t, {
    async emails() { throw new ApiError("FOLDER_NOT_FOUND", "Folder doesn't exist.", 404); },
  });
  const result = await request("/outlook-folder/count=3?folder=Missing");
  assert.equal(result.status, 404);
  assert.deepEqual(await result.json(), { error: "FOLDER_NOT_FOUND", message: "Folder doesn't exist." });
});
