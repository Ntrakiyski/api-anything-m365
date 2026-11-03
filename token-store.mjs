import { randomBytes } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { STATE_DIR, sessions } from "./api-anything.mjs";

export function localToken() {
  const file = join(STATE_DIR, "api-token");
  if (!existsSync(file)) {
    sessions.writePrivate(file, randomBytes(32).toString("hex"));
  }
  const token = readFileSync(file, "utf8").trim();
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error("The local API token file is invalid.");
  return token;
}
