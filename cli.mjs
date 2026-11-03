import { FolderReader, ApiError } from "./folder.mjs";
import { browser, createSendRead } from "./api-anything.mjs";

try {
  const [operation, ...args] = process.argv.slice(2);
  const values = {};
  for (const argument of args) {
    const match = /^(folder|count)=(.*)$/.exec(argument);
    if (!match || Object.hasOwn(values, match[1])) {
      throw new ApiError("INVALID_REQUEST", "Supply folder and count exactly once.", 400);
    }
    values[match[1]] = match[2];
  }
  if (!["list-emails", "read-emails"].includes(operation) ||
      !Object.hasOwn(values, "folder") || !Object.hasOwn(values, "count")) {
    throw new ApiError("INVALID_REQUEST", "Usage: node cli.mjs list-emails|read-emails folder=NAME count=3", 400);
  }
  const result = await new FolderReader(createSendRead())
    .emails(values.folder, values.count, operation === "read-emails");
  console.log(JSON.stringify(result, null, 2));
} catch (error) {
  console.error(JSON.stringify({
    error: error instanceof ApiError ? error.code : "INTERNAL_ERROR",
    message: error instanceof ApiError ? error.message : "The request failed unexpectedly.",
  }));
  process.exitCode = 1;
} finally {
  await browser.closeBrowser();
}
