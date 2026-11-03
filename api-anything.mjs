import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { getCACertificates, setDefaultCACertificates } from "node:tls";
import { ApiError, assertReadRequest, OUTLOOK_ORIGIN, SITE } from "./folder.mjs";

export const STATE_DIR = fileURLToPath(new URL(".api-anything", import.meta.url));
process.env.API_ANYTHING_HOME = STATE_DIR;
setDefaultCACertificates([...getCACertificates("default"), ...getCACertificates("system")]);

// API Anything exposes no public no-heal transport, so this adapter pins its internal 0.3.2 modules.
const moduleUrl = name => new URL(name, import.meta.resolve("api-anything"));
export const browser = await import(moduleUrl("browser.js"));
export const sessions = await import(moduleUrl("session.js"));
const { send } = await import(moduleUrl("http.js"));
sessions.writePrivate(join(STATE_DIR, "settings.json"), JSON.stringify({ share: false }));

export function createSendRead() {
  return async (action, payload, context) => {
    assertReadRequest(action, payload, context);
    const session = sessions.loadSession(SITE);
    if (!session.values.authorization) {
      throw new ApiError("LOGIN_REQUIRED", "Run npm run login in the API project folder.", 401);
    }
    const url = `${OUTLOOK_ORIGIN}/owa/service.svc?action=${action}&app=Mail`;
    const body = JSON.stringify(payload);
    const operation = {
      name: action, readOnly: true, params: [],
      request: {
        method: "POST", url, body,
        headers: { "content-type": "application/json; charset=utf-8", action },
      },
      slots: [
        { ref: "session:authorization", at: ["header:authorization"] },
        ...(session.values.anchor ? [{ ref: "session:anchor", at: ["header:x-anchormailbox"] }] : []),
      ],
    };
    let response;
    try {
      response = await send(operation, {}, session, {
        site: SITE, minIntervalMs: 300, timeoutMs: 30000,
        fetchImpl(target, options) {
          if (target !== url || options.method !== "POST" || options.body !== body ||
              options.headers.action !== action || options.redirect !== "manual") {
            throw new ApiError("REQUEST_REFUSED", "An unexpected request or redirect was blocked.", 403);
          }
          return fetch(target, options);
        },
      });
    } catch (error) {
      if (error instanceof ApiError) throw error;
      throw new ApiError("OUTLOOK_UNREACHABLE", `${action} could not reach Outlook. Check network access and retry.`);
    }
    if ([401, 403, 440].includes(response.status)) {
      throw new ApiError("LOGIN_REQUIRED", "Outlook login expired. Run npm run login.", 401);
    }
    if (response.status === 429) {
      throw new ApiError("RATE_LIMITED", "Outlook rate-limited the read. Wait before retrying.", 429);
    }
    if (response.status !== 200 || !/json/i.test(response.headers["content-type"] ?? "")) {
      throw new ApiError("OUTLOOK_ERROR", `${action} returned HTTP ${response.status} instead of JSON.`);
    }
    if (response.setCookies?.length) {
      sessions.mergeCapture(SITE, response.setCookies);
    }
    try {
      return JSON.parse(response.body);
    } catch {
      throw new ApiError("INVALID_OUTLOOK_RESPONSE", `${action} returned invalid JSON.`);
    }
  };
}

export function hasSession() {
  return existsSync(join(STATE_DIR, "sessions", `${SITE}.json`)) &&
    Boolean(sessions.loadSession(SITE).values.authorization);
}
