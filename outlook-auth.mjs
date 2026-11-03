import { join } from "node:path";
import { ApiError, MAIL_URL, OUTLOOK_ORIGIN, SITE, folderRequest, responseMessages } from "./folder.mjs";
import { STATE_DIR, browser, createSendRead, sessions } from "./api-anything.mjs";

const profileDir = join(STATE_DIR, "profile");

function inspectRequest(request) {
  const url = new URL(request.url);
  if (url.origin === OUTLOOK_ORIGIN && url.pathname === "/owa/startupdata.ashx" &&
      request.headers.authorization?.startsWith("Bearer ")) {
    const values = { authorization: request.headers.authorization };
    if (request.headers["x-anchormailbox"]) values.anchor = request.headers["x-anchormailbox"];
    return { from: request.url, values };
  }
}

function blockRequest(request) {
  const url = new URL(request.url);
  if (url.origin === OUTLOOK_ORIGIN &&
      ["/owa/startupdata.ashx", "/owa/service.svc"].includes(url.pathname)) return true;
  if (["GET", "HEAD", "OPTIONS"].includes(request.method)) return false;
  return url.protocol !== "https:" || url.hostname !== "login.microsoftonline.com";
}

export async function verifyOutlookSession() {
  const folders = responseMessages(await createSendRead()("GetFolder", folderRequest()), "GetFolder")
    .flatMap(result => result.Folders ?? []);
  if (folders.length !== 1 || typeof folders[0].FolderId?.Id !== "string") {
    throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Sign-in did not resolve to an Outlook mailbox.");
  }
}

export async function authenticateOutlook({ interactive = false, signal, onProgress = () => {} } = {}) {
  let closing;
  const abort = () => { closing = browser.closeBrowser(); };
  signal?.addEventListener("abort", abort, { once: true });
  try {
    signal?.throwIfAborted();
    let captured;
    let cookies;
    if (interactive) {
      const context = await browser.openBrowser({ profileDir, headless: false });
      signal?.throwIfAborted();
      await context.route("**/*", route => {
        const request = route.request();
        const value = { method: request.method(), url: request.url(), headers: request.headers() };
        captured = inspectRequest(value) ?? captured;
        return blockRequest(value) ? route.abort() : route.continue();
      });
      await context.routeWebSocket(/.*/, socket => socket.close());
      onProgress("Complete Microsoft sign-in, then close that sign-in window. Mail stays on a loading screen because mailbox data requests are blocked.");
      const result = await browser.login({ url: MAIL_URL, profileDir, waitForEnter: false });
      cookies = result.closed;
    }
    signal?.throwIfAborted();
    if (!captured) {
      const result = await browser.runTrigger({
        url: MAIL_URL, profileDir, headless: true, timeoutMs: 45000,
        steps: [{ action: "wait", ms: 6000 }],
        intercept(exchange) {
          captured = inspectRequest(exchange.request) ?? captured;
          return exchange.resourceType === "websocket" || blockRequest(exchange.request);
        },
      });
      cookies = result.cookies;
    }
    signal?.throwIfAborted();
    if (!captured || !Array.isArray(cookies)) {
      throw new ApiError("LOGIN_REQUIRED", "Microsoft sign-in is required. Connect again and complete sign-in.", 401);
    }
    const session = sessions.mergeCapture(SITE,
      cookies.filter(cookie => ["outlook.cloud.microsoft", ".outlook.cloud.microsoft"].includes(cookie.domain)),
      captured, { settings: [], urls: [`${OUTLOOK_ORIGIN}/owa/service.svc`] });
    sessions.saveSession(SITE, { ...session, source: "window" });
    await verifyOutlookSession();
    signal?.throwIfAborted();
  } catch (error) {
    if (signal?.aborted) throw signal.reason;
    if (error instanceof ApiError) throw error;
    console.error(`Microsoft connection failed: ${error.name ?? "Error"}`);
    throw new ApiError("LOGIN_FAILED", "Could not complete Microsoft sign-in. Check the sign-in window and try again.");
  } finally {
    signal?.removeEventListener("abort", abort);
    if (closing) await closing;
    await browser.closeBrowser();
  }
}
