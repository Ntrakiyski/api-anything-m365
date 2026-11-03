# Outlook folder read-only API

A local API backed by **API Anything 0.3.2**. Each request selects an Outlook
folder by display name and chooses whether to include full message bodies.
No Orca, Microsoft Graph integration, cloud proxy, catalog sharing, or
agent-configuration changes are used.

## Run

Requires Node.js 24 and an installed Chromium browser. API Anything's browser is
used only for local sign-in; API calls use its HTTP transport.

```powershell
# Run from the hammer-hai-api project folder.
npm ci --ignore-scripts
npm start
```

Open **http://127.0.0.1:3838/** in your usual browser. Select **Connect Microsoft**,
enter a folder's display name and an email count, choose whether to include full
bodies, then select **Read emails**. Opening the page checks the saved connection
using mailbox metadata; it never loads emails automatically.

The UI is plain HTML, CSS, and JavaScript served by this Node backend. No Vite
or build step is needed, but opening the HTML file directly will not work.
Each person runs their own local copy and signs into their own Microsoft account.
This is not a shared, hosted, multi-account service.

Connect first checks the saved credentials, then tries a headless refresh if a
browser profile exists. A valid saved login needs **no new visible window**.
First-time sign-in or new Microsoft MFA needs API Anything's separate isolated
browser window; this integration cannot sign in inside the HTML page or your
usual browser tab. Complete Microsoft sign-in, then **close the sign-in window**.
The mail loading screen may remain: mailbox data requests are blocked during
sign-in. Only a mailbox-root `GetFolder` metadata request verifies the login.
The app shows connection progress and provides **Cancel sign-in**.

The server listens on **127.0.0.1:3838**, never on external interfaces. Set `PORT`
before starting to choose another port.

## Requests

| Method | Endpoint | Returns |
|---|---|---|
| GET | `/outlook-folder/count=3?folder=HAMMER-HAI&includeBody=true` | Newest emails with full bodies, To, and CC |
| GET | `/outlook-folder/count=3?folder=HAMMER-HAI&includeBody=false` | Newest emails' metadata and previews only |
| GET | `/health` | Server liveness only; not a claim that Outlook authentication is valid |

External API clients use a local bearer token. Read it locally using
`npm run token --silent`; do not share it. Example PowerShell requests:

```powershell
$token = (npm run token --silent).Trim()
$headers = @{ Authorization = "Bearer $token" }
Invoke-RestMethod 'http://127.0.0.1:3838/outlook-folder/count=3?folder=HAMMER-HAI&includeBody=true' -Headers $headers
Invoke-RestMethod 'http://127.0.0.1:3838/outlook-folder/count=3?folder=Inbox&includeBody=false' -Headers $headers
```

Or use curl, with URL encoding for folder names containing spaces or symbols:

```powershell
curl.exe --get 'http://127.0.0.1:3838/outlook-folder/count=3' `
  -H "Authorization: Bearer $token" `
  --data-urlencode 'folder=Project A' `
  --data-urlencode 'includeBody=true'
```

| Input | Location | Rules |
|---|---|---|
| `count` | Path, after `count=` | Required positive decimal integer, at most `9007199254740991`; no leading zeros |
| `folder` | Query | Required non-empty display name; case-insensitive, with surrounding whitespace trimmed |
| `includeBody` | Query | Optional literal `true` or `false`; defaults to `true` |

Folder discovery searches mailbox folder metadata, including nested folders,
not email contents. A name must identify exactly one folder. Duplicate display
names return `409 AMBIGUOUS_FOLDER` instead of choosing one arbitrarily.
Names are literal display names, not folder paths or IDs. A non-email folder
returns `400 NOT_MAIL_FOLDER`.

Emails are ordered by received time, newest first. The reader pages through
the selected folder only, at most 100 items
per Outlook request, and returns at most `count` individual emails, **not
conversation threads**. If the folder has fewer, all available emails are
returned with the actual `returnedCount`. There is no arbitrary total-count
cap beyond JavaScript's safe integer limit.

The response shape is:

```json
{
  "folder": "HAMMER-HAI",
  "requestedCount": 1,
  "returnedCount": 1,
  "newestFirst": true,
  "hasMore": true,
  "emails": [
    {
      "id": "<Outlook item ID>",
      "subject": "<subject>",
      "from": { "name": "<sender>", "email": "<address>" },
      "receivedAt": "<ISO date>",
      "preview": "<preview>",
      "isRead": false,
      "hasAttachments": false,
      "body": { "contentType": "Text", "content": "<complete body>" },
      "to": [],
      "cc": []
    }
  ]
}
```

The example shows the response for `count=1`. `folder` contains the resolved
display name, and `returnedCount` equals the number of email entries.
`body`, `to`, and `cc` appear only when `includeBody=true`.
Bodies are returned as data, never executed. Attachments, remote images, and
links are not fetched. Encrypted or otherwise unavailable bodies fail explicitly;
the API does not silently substitute previews.

CLI equivalents, without starting the HTTP server:

```powershell
npm run list -- folder=HAMMER-HAI count=3
npm run read -- "folder=Project A" count=3
```

For a nonexistent folder, HTTP status **404** is returned with:

```json
{
  "error": "FOLDER_NOT_FOUND",
  "message": "Folder doesn't exist."
}
```

An existing empty folder returns `200` with `returnedCount: 0` and `emails: []`.
Missing/invalid parameters return `400`; invalid or expired authentication
returns `401`. The old `/hammer-hai/list-emails` and `/hammer-hai/read-emails`
routes have been replaced by the single `/outlook-folder/count=...` route.

## Folder and read-only restrictions

- The requested name is resolved to a folder ID using read-only metadata.
  That ID is fixed for the duration of this request, not shared across requests.
  Folder-ID, message-ID, action, field, sort, URL, and cursor overrides are rejected.
- Only exact, locally constructed `FindFolder`, `GetFolder`, `FindItem`, and `GetItem` POST
  templates are allowed. POST here fetches data; it does not mean a mailbox edit.
  Email endpoints accept GET only. Local UI sign-in controls use POST, but do
  not send, edit, or mark emails.
- Email lists use shallow traversal of the selected folder, excluding its subfolders.
  Each returned item's parent
  folder is checked before use. Body requests accept only IDs from the current
  verified listing, and their returned parent folder is checked again.
- There are no send, reply, edit, delete, move, mark-read, or mark-unread actions.
  The existing read/unread status is preserved.
- No browser fallback, automatic healing, registry lookup, or redirects to
  another endpoint are allowed during API calls. Unexpected requests fail closed.
- A renamed folder during lookup, a moved/deleted item, invalid paging, or an unexpected
  response fails explicitly instead of returning another folder or partial data.
  Concurrent mailbox changes may require retrying a request.

**This is application-enforced folder isolation, not a folder-scoped Microsoft
permission.** The Outlook sign-in remains mailbox-wide. For body reads, folder
membership is verified immediately before and after the item lookup; Outlook's
item-ID API does not provide an atomic folder-membership condition, so a
concurrent move can be detected and rejected but cannot be prevented here.

## Authentication and local data

API Anything's cookies, bearer token, browser profile, and local API token remain
in `.api-anything`, excluded by `.gitignore`. Sharing is forced off. Email results
and raw captures are not written to disk by this application. The browser may
keep its ordinary authentication/cache files; protect this directory with your
Windows account permissions. Do not copy it into source control or publish it.
When sharing the project, share only its source files and dependency manifests,
never `.api-anything` or `node_modules`.

The UI uses a temporary HttpOnly, SameSite=Strict cookie instead of exposing the
local API bearer token or Microsoft credentials to JavaScript. Login actions
also require a same-origin request and a CSRF token. Browser sessions expire
after eight hours of inactivity or a server restart; reload the page to renew
the local browser session. **Clear results** removes returned messages from the
page and its in-memory state, not from Outlook. Email HTML is displayed as inert
text, with no executable markup, remote images, or links.

Outlook tokens expire. In the UI, select **Connect Microsoft** again.
For terminal-only sign-in, use `npm run login`. A `401 LOGIN_REQUIRED` from an
external API client can also be handled with:

```powershell
npm run refresh-login
```

This refreshes using API Anything's saved browser session while blocking mailbox
data requests. If Microsoft requires interactive authentication, run
`npm run login` instead. API calls never silently launch a browser or broaden access.
The local API token is independent of the Outlook login.

The adapter intentionally uses API Anything's pinned internal HTTP/browser
modules because its public `call()` can heal and run a browser fallback. Do not
upgrade the dependency without re-verifying the exact request guards.

## Checks

```powershell
npm test
```

Tests use synthetic email data, not the live mailbox.
