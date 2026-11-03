import { randomBytes, timingSafeEqual } from "node:crypto";
import { ApiError } from "./folder.mjs";

const COOKIE_NAME = "outlook_ui";
const TTL_MS = 8 * 60 * 60 * 1000;
const equal = (actual, expected) => {
  if (typeof actual !== "string") return false;
  const left = Buffer.from(actual);
  const right = Buffer.from(expected);
  return left.length === right.length && timingSafeEqual(left, right);
};

export class WebSessions {
  constructor() {
    this.sessions = new Map();
  }

  get(request) {
    const values = (request.headers.cookie ?? "").split(";")
      .map(value => value.trim()).filter(value => value.startsWith(`${COOKIE_NAME}=`));
    if (values.length !== 1) return undefined;
    const id = values[0].slice(COOKIE_NAME.length + 1);
    const session = this.sessions.get(id);
    if (!session) return undefined;
    if (session.expires <= Date.now()) {
      this.sessions.delete(id);
      return undefined;
    }
    session.expires = Date.now() + TTL_MS;
    return session;
  }

  issue(request, response) {
    const existing = this.get(request);
    if (existing) return existing;
    for (const [id, session] of this.sessions) {
      if (session.expires <= Date.now()) this.sessions.delete(id);
    }
    if (this.sessions.size >= 64) this.sessions.delete(this.sessions.keys().next().value);
    const id = randomBytes(32).toString("hex");
    const session = { csrf: randomBytes(32).toString("hex"), expires: Date.now() + TTL_MS };
    this.sessions.set(id, session);
    response.setHeader("set-cookie", `${COOKIE_NAME}=${id}; HttpOnly; SameSite=Strict; Path=/`);
    return session;
  }

  require(request) {
    const session = this.get(request);
    if (!session) throw new ApiError("UI_SESSION_REQUIRED", "Reload the app to start a local browser session.", 401);
    return session;
  }

  requireAction(request, origin) {
    const session = this.require(request);
    if (request.headers.origin !== origin || !equal(request.headers["x-csrf-token"], session.csrf)) {
      throw new ApiError("REQUEST_REFUSED", "This action must come from the local app.", 403);
    }
    return session;
  }
}
