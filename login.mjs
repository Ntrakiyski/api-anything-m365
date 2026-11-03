import { ApiError } from "./folder.mjs";
import { authenticateOutlook } from "./outlook-auth.mjs";

const args = process.argv.slice(2);
const refresh = args.length === 1 && args[0] === "--refresh";

try {
  if (args.length && !refresh) {
    throw new ApiError("INVALID_REQUEST", "Usage: node login.mjs [--refresh]", 400);
  }
  await authenticateOutlook({ interactive: !refresh, onProgress: message => console.log(message) });
  console.log("Signed in. Outlook mailbox verified; no emails fetched during sign-in.");
} catch (error) {
  console.error(error instanceof ApiError ? `${error.code}: ${error.message}` :
    `LOGIN_FAILED: API Anything could not complete sign-in (${error.name ?? "Error"}).`);
  process.exitCode = 1;
}
