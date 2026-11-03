import test from "node:test";
import assert from "node:assert/strict";
import { OutlookConnection } from "./connection.mjs";
import { ApiError } from "./folder.mjs";

const expired = () => { throw new ApiError("LOGIN_REQUIRED", "Sign-in required.", 401); };
function create(overrides = {}) {
  const calls = [];
  const connection = new OutlookConnection({
    verify: async () => { calls.push("verify"); },
    refresh: async () => { calls.push("refresh"); },
    signIn: async () => { calls.push("signIn"); },
    canRefresh: () => true,
    ...overrides,
  });
  return { connection, calls };
}

test("page startup verifies saved credentials once without opening any browser", async () => {
  const { connection, calls } = create();
  assert.equal(connection.snapshot().status, "checking");
  await connection.pending;
  assert.equal(connection.snapshot().canRead, true);
  connection.snapshot();
  assert.deepEqual(calls, ["verify"]);
});

test("missing saved credentials are reported disconnected, not silently connected", async () => {
  const { connection } = create({ verify: expired });
  connection.snapshot();
  await connection.pending;
  assert.equal(connection.snapshot().status, "disconnected");
  assert.equal(connection.snapshot().canRead, false);
});

test("Connect reuses an already valid saved login without launching a browser", async () => {
  const { connection, calls } = create();
  connection.start();
  await connection.pending;
  assert.equal(connection.snapshot().status, "connected");
  assert.deepEqual(calls, ["verify"]);
});

test("expired login attempts silent refresh before asking for interactive sign-in", async () => {
  const { connection, calls } = create({ verify: expired });
  connection.start();
  await connection.pending;
  assert.equal(connection.snapshot().canRead, true);
  assert.deepEqual(calls, ["refresh"]);
});

test("interactive sign-in is used only when silent refresh reports authentication is needed", async () => {
  const { connection, calls } = create({ verify: expired, refresh: expired });
  connection.start();
  await connection.pending;
  assert.equal(connection.snapshot().status, "connected");
  assert.deepEqual(calls, ["signIn"]);
});

test("first use skips silent refresh when no browser profile exists", async () => {
  const { connection, calls } = create({ verify: expired, canRefresh: () => false });
  connection.start();
  await connection.pending;
  assert.deepEqual(calls, ["signIn"]);
});

test("network failure does not get hidden by opening an interactive login window", async () => {
  const { connection, calls } = create({
    verify: async () => { throw new ApiError("OUTLOOK_UNREACHABLE", "Outlook cannot be reached."); },
  });
  connection.start();
  await connection.pending;
  assert.equal(connection.snapshot().status, "error");
  assert.equal(connection.snapshot().error, "OUTLOOK_UNREACHABLE");
  assert.equal(connection.snapshot().canRead, false);
  assert.deepEqual(calls, []);
});

test("cancelling interactive sign-in propagates an abort and leaves the UI disconnected", async () => {
  let opened;
  const visible = new Promise(resolve => { opened = resolve; });
  const { connection } = create({
    verify: expired, canRefresh: () => false,
    signIn: signal => new Promise((resolve, reject) => {
      opened();
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    }),
  });
  connection.start();
  await visible;
  assert.equal(connection.snapshot().status, "signing_in");
  assert.throws(() => connection.start(), error => error.code === "AUTH_BUSY");
  const pending = connection.pending;
  assert.equal(connection.cancel().status, "cancelling");
  await pending;
  assert.equal(connection.snapshot().status, "disconnected");
  assert.equal(connection.isConnecting, false);
  assert.throws(() => connection.cancel(), error => error.code === "NO_SIGN_IN");
});

test("a sign-in timeout is explicit and does not leave a job active", async () => {
  const { connection } = create({
    timeoutMs: 15, verify: expired, canRefresh: () => false,
    signIn: signal => new Promise((resolve, reject) => {
      signal.addEventListener("abort", () => reject(signal.reason), { once: true });
    }),
  });
  connection.start();
  await Promise.all([connection.pending, new Promise(resolve => setTimeout(resolve, 25))]);
  assert.equal(connection.snapshot().status, "error");
  assert.equal(connection.snapshot().error, "AUTH_TIMEOUT");
  assert.equal(connection.isConnecting, false);
});

test("authentication errors from a read invalidate the displayed connection", async () => {
  const { connection } = create();
  connection.start();
  await connection.pending;
  connection.requireLogin();
  assert.equal(connection.snapshot().canRead, false);
  assert.equal(connection.snapshot().status, "disconnected");
});

test("a late success after cancellation cannot mark the connection as ready", async () => {
  let finish;
  const { connection } = create({ verify: () => new Promise(resolve => { finish = resolve; }) });
  connection.start();
  const pending = connection.pending;
  connection.cancel();
  finish();
  await pending;
  assert.equal(connection.snapshot().status, "disconnected");
  assert.equal(connection.snapshot().canRead, false);
});
