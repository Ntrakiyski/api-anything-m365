import test from "node:test";
import assert from "node:assert/strict";
import { WebSessions } from "./web-session.mjs";

function issue(sessions, request = { headers: {} }) {
  let cookie;
  const session = sessions.issue(request, { setHeader: (name, value) => { cookie = value.split(";")[0]; } });
  return { session, request: { headers: { cookie } } };
}

test("browser sessions are unpredictable, reused on reload, and reject forged or duplicate cookies", () => {
  const sessions = new WebSessions();
  const first = issue(sessions);
  const second = issue(sessions);
  assert.notEqual(first.session.csrf, second.session.csrf);
  assert.notEqual(first.request.headers.cookie, second.request.headers.cookie);
  assert.equal(sessions.get(first.request), first.session);
  assert.equal(sessions.issue(first.request, { setHeader: () => assert.fail("Unexpected session rotation") }), first.session);
  assert.equal(sessions.get({ headers: { cookie: "outlook_ui=forged" } }), undefined);
  assert.equal(sessions.get({ headers: { cookie: `${first.request.headers.cookie}; ${second.request.headers.cookie}` } }), undefined);
});

test("expired sessions are removed and cannot authorize reads or actions", () => {
  const sessions = new WebSessions();
  const { session, request } = issue(sessions);
  session.expires = Date.now() - 1;
  assert.equal(sessions.get(request), undefined);
  assert.equal(sessions.sessions.size, 0);
  assert.throws(() => sessions.require(request), error => error.code === "UI_SESSION_REQUIRED");
});

test("browser session storage is capped and expired sessions are pruned on creation", () => {
  const sessions = new WebSessions();
  const first = issue(sessions);
  for (let index = 0; index < 64; index++) issue(sessions);
  assert.equal(sessions.sessions.size, 64);
  assert.equal(sessions.get(first.request), undefined);
  for (const session of sessions.sessions.values()) session.expires = Date.now() - 1;
  issue(sessions);
  assert.equal(sessions.sessions.size, 1);
});
