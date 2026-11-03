import { createServer } from "node:http";
import { timingSafeEqual } from "node:crypto";
import { readFileSync } from "node:fs";
import { ApiError, parseCount, parseFolderName, parseIncludeBody } from "./folder.mjs";
import { WebSessions } from "./web-session.mjs";

const authorized = (value, token) => {
  if (typeof value !== "string" || !value.startsWith("Bearer ")) return false;
  const actual = Buffer.from(value.slice(7));
  const expected = Buffer.from(token);
  return actual.length === expected.length && timingSafeEqual(actual, expected);
};

function reply(response, status, value, headers = {}) {
  response.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    "x-content-type-options": "nosniff",
    "content-security-policy": "default-src 'none'; frame-ancestors 'none'",
    ...headers,
  });
  response.end(JSON.stringify(value));
}

export function loadWebAssets() {
  return new Map([
    ["/", { type: "text/html; charset=utf-8", body: readFileSync(new URL("public/index.html", import.meta.url)) }],
    ["/app.css", { type: "text/css; charset=utf-8", body: readFileSync(new URL("public/app.css", import.meta.url)) }],
    ["/app.js", { type: "text/javascript; charset=utf-8", body: readFileSync(new URL("public/app.js", import.meta.url)) }],
  ]);
}

const APP_CSP = "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'";

async function requireEmptyJson(request) {
  if (!/^application\/json(?:\s*;|$)/i.test(request.headers["content-type"] ?? "")) {
    throw new ApiError("INVALID_REQUEST", "This action requires an empty JSON object.", 400);
  }
  if (Number(request.headers["content-length"] ?? 0) > 256) {
    request.resume();
    throw new ApiError("INVALID_REQUEST", "The action body is too large.", 413);
  }
  let length = 0;
  const chunks = [];
  for await (const chunk of request) {
    length += chunk.length;
    if (length <= 256) chunks.push(chunk);
  }
  if (length > 256) throw new ApiError("INVALID_REQUEST", "The action body is too large.", 413);
  let value;
  try {
    value = JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new ApiError("INVALID_REQUEST", "This action requires valid JSON.", 400);
  }
  if (!value || Array.isArray(value) || typeof value !== "object" || Object.keys(value).length !== 0) {
    throw new ApiError("INVALID_REQUEST", "This action accepts only an empty JSON object.", 400);
  }
}

export function createApiServer(reader, token, { connection = null, assets = null } = {}) {
  if (typeof token !== "string" || !/^[a-f0-9]{64}$/.test(token)) {
    throw new Error("A 32-byte local API token is required.");
  }
  const webSessions = new WebSessions();
  let busy = false;
  const server = createServer({ maxHeaderSize: 16384 }, async (request, response) => {
    try {
      const origin = `http://${request.headers.host}`;
      const openingApp = Boolean(assets) && request.method === "GET" && request.url === "/" &&
        request.headers["sec-fetch-mode"] === "navigate" && request.headers["sec-fetch-dest"] === "document";
      if (!/^127\.0\.0\.1:\d+$/.test(request.headers.host ?? "") ||
          (request.headers.origin && request.headers.origin !== origin) ||
          (request.headers["sec-fetch-site"] === "cross-site" && !openingApp)) {
        throw new ApiError("LOCAL_REQUEST_REQUIRED", "Use a local API client on 127.0.0.1.", 403);
      }
      if (!request.url.startsWith("/") || request.url.startsWith("//")) {
        throw new ApiError("INVALID_REQUEST", "Absolute request URLs are not accepted.", 400);
      }
      const url = new URL(request.url, "http://127.0.0.1");
      if (url.origin !== "http://127.0.0.1") {
        throw new ApiError("INVALID_REQUEST", "The request URL must remain local.", 400);
      }
      if (request.method === "GET" && (request.headers["transfer-encoding"] ||
          (request.headers["content-length"] && request.headers["content-length"] !== "0"))) {
        throw new ApiError("INVALID_REQUEST", "GET requests must not contain a body.", 400);
      }
      const asset = assets?.get(url.pathname);
      if (asset && !url.search) {
        if (request.method !== "GET") {
          reply(response, 405, { error: "METHOD_NOT_ALLOWED", message: "Only GET is allowed." }, { allow: "GET" });
          return;
        }
        if (url.pathname === "/") webSessions.issue(request, response);
        response.writeHead(200, {
          "content-type": asset.type, "cache-control": "no-store",
          "x-content-type-options": "nosniff", "content-security-policy": APP_CSP,
          "referrer-policy": "no-referrer",
        });
        response.end(asset.body);
        return;
      }
      if (connection && ["/ui/status", "/ui/login", "/ui/login/cancel"].includes(url.pathname) && !url.search) {
        if (url.pathname === "/ui/status") {
          if (request.method !== "GET") {
            reply(response, 405, { error: "METHOD_NOT_ALLOWED", message: "Only GET is allowed." }, { allow: "GET" });
            return;
          }
          const session = webSessions.require(request);
          reply(response, 200, { ...connection.snapshot(), csrfToken: session.csrf });
          return;
        }
        if (request.method !== "POST") {
          reply(response, 405, { error: "METHOD_NOT_ALLOWED", message: "Only POST is allowed." }, { allow: "POST" });
          return;
        }
        webSessions.requireAction(request, origin);
        await requireEmptyJson(request);
        if (busy) throw new ApiError("READ_IN_PROGRESS", "Wait for the current email read before connecting.", 409);
        const status = url.pathname === "/ui/login" ? connection.start() : connection.cancel();
        reply(response, 202, status);
        return;
      }
      if (request.method !== "GET") {
        reply(response, 405, { error: "METHOD_NOT_ALLOWED", message: "Only GET is allowed." }, { allow: "GET" });
        return;
      }
      if (url.pathname === "/health" && !url.search) {
        reply(response, 200, { ok: true });
        return;
      }
      const allowed = request.headers.authorization !== undefined
        ? authorized(request.headers.authorization, token) : Boolean(webSessions.get(request));
      if (!allowed) {
        reply(response, 401, { error: "UNAUTHORIZED", message: "A local API bearer token is required." },
          { "www-authenticate": "Bearer" });
        return;
      }
      const match = /^\/outlook-folder\/count=([^/]*)$/.exec(url.pathname);
      if (!match) {
        throw new ApiError("NOT_FOUND", "Unknown API endpoint.", 404);
      }
      const entries = [...url.searchParams];
      if (entries.some(([key]) => !["folder", "includeBody"].includes(key)) ||
          url.searchParams.getAll("folder").length !== 1 ||
          url.searchParams.getAll("includeBody").length > 1) {
        throw new ApiError("INVALID_REQUEST",
          "Supply folder once and optionally includeBody=true or false. No other query parameters are allowed.", 400);
      }
      const count = parseCount(match[1]);
      const folder = parseFolderName(url.searchParams.get("folder"));
      const includeBody = parseIncludeBody(url.searchParams.get("includeBody") ?? undefined);
      if (connection?.isConnecting) {
        throw new ApiError("AUTH_IN_PROGRESS", "Finish connecting Microsoft before reading a folder.", 409);
      }
      if (busy) {
        reply(response, 429, { error: "BUSY", message: "Another email read is in progress. Retry after it completes." },
          { "retry-after": "5" });
        return;
      }
      busy = true;
      try {
        const result = await reader.emails(folder, count, includeBody);
        reply(response, 200, result);
      } finally {
        busy = false;
      }
    } catch (error) {
      const known = error instanceof ApiError;
      if (known && error.code === "LOGIN_REQUIRED") connection?.requireLogin();
      if (!known || error.status >= 500) {
        console.error(`Email API request failed: ${known ? error.code : "INTERNAL_ERROR"}`);
      }
      reply(response, known ? error.status : 500, {
        error: known ? error.code : "INTERNAL_ERROR",
        message: known ? error.message : "The request failed unexpectedly.",
      });
    }
  });
  server.requestTimeout = 15000;
  server.headersTimeout = 10000;
  server.keepAliveTimeout = 5000;
  return server;
}
