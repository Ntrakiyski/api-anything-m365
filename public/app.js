const byId = id => document.getElementById(id);
const elements = Object.fromEntries([
  "connection-badge", "connection-message", "connect-button", "cancel-button",
  "reader-form", "request-fields", "folder", "count", "include-body", "read-button",
  "read-button-label", "request-hint", "error-notice", "error-title", "error-message",
  "results", "results-title", "results-summary", "clear-button", "loading-state",
  "empty-state", "empty-title", "empty-message", "mail-workspace", "email-list",
  "email-subject", "email-date", "email-read-status", "email-from", "email-to",
  "email-cc", "email-to-row", "email-cc-row", "email-body", "email-body-label",
  "email-attachments", "preview-note", "json-panel", "json-response",
].map(id => [id, byId(id)]));

let connection = { status: "checking", canRead: false, busy: true };
let csrfToken;
let reading = false;
let needsReload = false;
let responseData;
let pollTimer;
const dateFormat = new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" });
const shortDateFormat = new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" });

async function api(url, options = {}) {
  let response;
  try {
    response = await fetch(url, { credentials: "same-origin", cache: "no-store", ...options });
  } catch {
    throw { code: "APP_UNREACHABLE", message: "The local app is not reachable. Make sure npm start is still running, then try again." };
  }
  let data;
  try {
    data = await response.json();
  } catch {
    throw { code: "INVALID_RESPONSE", message: "The local app returned an unexpected response. Reload the page and try again." };
  }
  if (!response.ok) throw { code: data.error ?? "REQUEST_FAILED", message: data.message ?? "The request failed." };
  return data;
}

function showError(error) {
  elements["error-title"].textContent = error.code === "FOLDER_NOT_FOUND" ? "Folder doesn't exist." : "Couldn't complete that request.";
  elements["error-message"].textContent = `${error.message} (${error.code})`;
  elements["error-notice"].hidden = false;
  elements["error-notice"].focus();
}

function clearError() {
  elements["error-notice"].hidden = true;
  elements["error-title"].textContent = "";
  elements["error-message"].textContent = "";
}

function updateControls() {
  elements["read-button"].disabled = reading || connection.busy || !connection.canRead;
  elements["read-button-label"].textContent = reading ? "Reading folder..." : "Read emails";
  elements["request-fields"].disabled = reading;
  elements["connect-button"].disabled = reading || connection.busy;
  elements["cancel-button"].disabled = connection.status === "cancelling";
  elements["cancel-button"].hidden = !["refreshing", "signing_in", "cancelling"].includes(connection.status);
  elements["connect-button"].textContent = needsReload ? "Reload app" :
    connection.canRead ? "Check connection" : "Connect Microsoft";
  elements["request-hint"].textContent = reading ? "Reading only the selected folder. Outlook will not be changed." :
    connection.canRead ? "Ready. No emails are loaded until you choose Read emails." :
      connection.busy ? "Waiting for your Microsoft connection." : "Connect Microsoft to read a folder.";
}

function renderConnection(data) {
  connection = data;
  if (data.csrfToken) csrfToken = data.csrfToken;
  const labels = {
    checking: "Checking", connected: "Connected", disconnected: "Not connected",
    refreshing: "Connecting", signing_in: "Sign-in needed", cancelling: "Cancelling", error: "Connection issue",
  };
  elements["connection-badge"].textContent = labels[data.status] ?? "Connection issue";
  elements["connection-badge"].dataset.status = data.status;
  elements["connection-message"].textContent = data.message;
  updateControls();
  clearTimeout(pollTimer);
  if (data.busy) pollTimer = setTimeout(loadStatus, 1200);
}

async function loadStatus() {
  try {
    renderConnection(await api("/ui/status"));
  } catch (error) {
    needsReload = error.code === "UI_SESSION_REQUIRED";
    if (needsReload) csrfToken = undefined;
    renderConnection({ status: "error", canRead: false, busy: false, message: error.message });
  }
}

async function connectAction(path) {
  if (needsReload) {
    window.location.reload();
    return;
  }
  clearError();
  if (!csrfToken) {
    await loadStatus();
    if (!csrfToken || connection.busy) return;
  }
  clearResults();
  elements["connect-button"].disabled = true;
  elements["cancel-button"].disabled = true;
  try {
    const data = await api(path, {
      method: "POST",
      headers: { "content-type": "application/json", "x-csrf-token": csrfToken },
      body: "{}",
    });
    renderConnection(data);
  } catch (error) {
    showError(error);
    await loadStatus();
  }
}

function clearResults() {
  responseData = undefined;
  elements["results-title"].textContent = "Your emails";
  elements["results-summary"].textContent = "Messages appear here only when you request them.";
  elements["mail-workspace"].hidden = true;
  elements["email-list"].replaceChildren();
  elements["json-panel"].hidden = true;
  elements["json-panel"].open = false;
  elements["json-response"].textContent = "";
  elements["email-body"].textContent = "";
  elements["email-subject"].textContent = "";
  for (const id of ["email-from", "email-to", "email-cc", "email-date", "email-read-status", "email-attachments"]) {
    elements[id].textContent = "";
  }
  elements["email-date"].removeAttribute("datetime");
  elements["clear-button"].hidden = true;
  elements["empty-state"].hidden = false;
  elements["empty-title"].textContent = "A window into one folder.";
  elements["empty-message"].textContent = "Connect Microsoft, enter a folder name, and choose Read emails. Nothing in Outlook will be changed.";
}

function mailboxText(value) {
  if (!value) return "Not provided";
  if (value.name && value.email) return `${value.name} <${value.email}>`;
  return value.email || value.name || "Not provided";
}

function renderEmail(index) {
  const email = responseData.emails[index];
  for (const [buttonIndex, button] of [...elements["email-list"].querySelectorAll("button")].entries()) {
    button.classList.toggle("is-selected", buttonIndex === index);
    button.setAttribute("aria-pressed", String(buttonIndex === index));
  }
  elements["email-subject"].textContent = email.subject || "(No subject)";
  elements["email-from"].textContent = mailboxText(email.from);
  elements["email-date"].textContent = dateFormat.format(new Date(email.receivedAt));
  elements["email-date"].dateTime = email.receivedAt;
  elements["email-read-status"].textContent = email.isRead ? "Read in Outlook" : "Unread in Outlook";
  elements["email-attachments"].textContent = email.hasAttachments ? "Has attachments (not fetched)" : "";
  const withBody = Object.hasOwn(email, "body");
  elements["email-to-row"].hidden = !withBody;
  elements["email-cc-row"].hidden = !withBody || !email.cc?.length;
  elements["email-to"].textContent = withBody ? (email.to?.map(mailboxText).join("; ") || "Not provided") : "";
  elements["email-cc"].textContent = withBody ? (email.cc?.map(mailboxText).join("; ") || "") : "";
  elements["email-body-label"].textContent = withBody ? "Full message" : "Preview only";
  // Email HTML is intentionally shown as inert text; no links or remote images are loaded.
  elements["email-body"].textContent = withBody ? email.body.content : email.preview || "No preview available.";
  elements["preview-note"].hidden = withBody;
}

function renderResults(data) {
  responseData = data;
  elements["results-title"].textContent = data.folder;
  elements["results-summary"].textContent =
    `${data.returnedCount} of ${data.requestedCount} requested ${data.requestedCount === 1 ? "email" : "emails"} \u00b7 Newest first${data.hasMore ? " \u00b7 More available" : ""}`;
  elements["clear-button"].hidden = false;
  elements["empty-state"].hidden = data.emails.length > 0;
  elements["mail-workspace"].hidden = data.emails.length === 0;
  elements["json-panel"].hidden = false;
  elements["json-response"].textContent = JSON.stringify(data, null, 2);
  if (data.emails.length === 0) {
    elements["empty-title"].textContent = "This folder is empty.";
    elements["empty-message"].textContent = "The folder exists, but Outlook returned no emails. Choose another folder to read.";
    return;
  }
  const list = data.emails.map((email, index) => {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    const meta = document.createElement("span");
    meta.className = "list-meta";
    const sender = document.createElement("span");
    sender.textContent = email.from?.name || email.from?.email || "Sender not provided";
    const date = document.createElement("span");
    date.textContent = shortDateFormat.format(new Date(email.receivedAt));
    meta.append(sender, date);
    const subject = document.createElement("span");
    subject.className = "list-subject";
    if (!email.isRead) {
      const dot = document.createElement("span");
      dot.className = "unread-dot";
      dot.setAttribute("aria-hidden", "true");
      subject.append(dot);
    }
    subject.append(document.createTextNode(email.subject || "(No subject)"));
    const preview = document.createElement("span");
    preview.className = "list-preview";
    preview.textContent = email.preview || "No preview available.";
    button.append(meta, subject, preview);
    button.addEventListener("click", () => renderEmail(index));
    item.append(button);
    return item;
  });
  elements["email-list"].replaceChildren(...list);
  renderEmail(0);
}

elements["reader-form"].addEventListener("submit", async event => {
  event.preventDefault();
  clearError();
  const folder = elements.folder.value.trim();
  const count = elements.count.value;
  if (!folder || !/^[1-9]\d*$/.test(count) || !Number.isSafeInteger(Number(count))) {
    showError({ code: "INVALID_REQUEST", message: "Enter a folder name and a positive whole-number email count." });
    return;
  }
  if (reading || !connection.canRead || connection.busy) return;
  clearResults();
  reading = true;
  updateControls();
  elements.results.setAttribute("aria-busy", "true");
  elements["loading-state"].hidden = false;
  elements["empty-state"].hidden = true;
  const query = new URLSearchParams({ folder, includeBody: String(elements["include-body"].checked) });
  try {
    const data = await api(`/outlook-folder/count=${count}?${query}`);
    renderResults(data);
    elements["results-title"].focus();
  } catch (error) {
    elements["empty-state"].hidden = false;
    showError(error);
    if (error.code === "LOGIN_REQUIRED" || error.code === "UNAUTHORIZED") await loadStatus();
  } finally {
    reading = false;
    elements.results.setAttribute("aria-busy", "false");
    elements["loading-state"].hidden = true;
    updateControls();
  }
});

elements["connect-button"].addEventListener("click", () => connectAction("/ui/login"));
elements["cancel-button"].addEventListener("click", () => connectAction("/ui/login/cancel"));
elements["clear-button"].addEventListener("click", () => { clearResults(); clearError(); elements.folder.focus(); });
window.addEventListener("pagehide", () => clearTimeout(pollTimer));
loadStatus();
