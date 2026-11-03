import { FolderReader } from "./folder.mjs";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { STATE_DIR, browser, createSendRead } from "./api-anything.mjs";
import { createApiServer, loadWebAssets } from "./http-api.mjs";
import { localToken } from "./token-store.mjs";
import { OutlookConnection } from "./connection.mjs";
import { authenticateOutlook, verifyOutlookSession } from "./outlook-auth.mjs";

const port = process.env.PORT ?? "3838";
if (!/^[1-9]\d*$/.test(port) || Number(port) > 65535) {
  throw new Error("PORT must be an integer between 1 and 65535.");
}
const connection = new OutlookConnection({
  verify: verifyOutlookSession,
  canRefresh: () => existsSync(join(STATE_DIR, "profile")),
  refresh: signal => authenticateOutlook({ signal }),
  signIn: signal => authenticateOutlook({ interactive: true, signal }),
});
const server = createApiServer(new FolderReader(createSendRead()), localToken(), {
  connection, assets: loadWebAssets(),
});
server.on("error", error => {
  console.error(`Email API server failed: ${error.code ?? "SERVER_ERROR"}`);
  process.exitCode = 1;
});
server.listen(Number(port), "127.0.0.1", () => {
  console.log(`Outlook folder read-only API: http://127.0.0.1:${port}`);
});
for (const signal of ["SIGINT", "SIGTERM"]) {
  process.once(signal, () => {
    if (connection.isConnecting) connection.cancel();
    server.close(async () => { await browser.closeBrowser(); });
  });
}
