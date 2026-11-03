import { isDeepStrictEqual } from "node:util";

export const OUTLOOK_ORIGIN = "https://outlook.cloud.microsoft";
export const MAIL_URL = `${OUTLOOK_ORIGIN}/mail/`;
// Keep the existing private session key so saved logins continue to work.
export const SITE = "hammer-hai";

export class ApiError extends Error {
  constructor(code, message, status = 502) {
    super(message);
    this.code = code;
    this.status = status;
  }
}

export function parseCount(value) {
  if ((typeof value !== "string" && typeof value !== "number") ||
      !/^[1-9]\d*$/.test(String(value)) || !Number.isSafeInteger(Number(value))) {
    throw new ApiError("INVALID_COUNT", "count must be a positive safe integer.", 400);
  }
  return Number(value);
}

export function parseFolderName(value) {
  if (typeof value !== "string" || !value.trim() || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new ApiError("INVALID_FOLDER", "folder must be a non-empty folder name without control characters.", 400);
  }
  return value.trim();
}

export function parseIncludeBody(value) {
  if (value === undefined || value === "true") return true;
  if (value === "false") return false;
  throw new ApiError("INVALID_INCLUDE_BODY", "includeBody must be true or false.", 400);
}

const nameKey = name => name.normalize("NFC").toLowerCase();
function assertFolderId(folderId) {
  if (typeof folderId !== "string" || !folderId) {
    throw new ApiError("REQUEST_REFUSED", "A resolved folder ID is required for email reads.", 403);
  }
}
function assertOffset(offset) {
  if (!Number.isSafeInteger(offset) || offset < 0) {
    throw new ApiError("INVALID_OFFSET", "Invalid internal paging offset.", 500);
  }
}
const property = FieldURI => ({ __type: "PropertyUri:#Exchange", FieldURI });
const envelope = (action, Body) => ({
  __type: `${action}JsonRequest:#Exchange`,
  Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
  Body: { __type: `${action}Request:#Exchange`, ...Body },
});
const folderReference = folderId => {
  assertFolderId(folderId);
  return { __type: "FolderId:#Exchange", Id: folderId };
};
const rootReference = () => ({ __type: "DistinguishedFolderId:#Exchange", Id: "msgfolderroot" });
const metadata = [
  "item:ParentFolderId", "item:Subject", "item:DateTimeReceived",
  "item:Preview", "item:HasAttachments", "message:From", "message:Sender", "message:IsRead",
];

export function folderRequest(folderId) {
  return envelope("GetFolder", {
    FolderShape: {
      __type: "FolderResponseShape:#Exchange", BaseShape: "Default",
      AdditionalProperties: [property("folder:FolderClass")],
    },
    FolderIds: [folderId === undefined ? rootReference() : folderReference(folderId)],
  });
}

export function findFoldersRequest(offset) {
  assertOffset(offset);
  return envelope("FindFolder", {
    FolderShape: { __type: "FolderResponseShape:#Exchange", BaseShape: "Default" },
    ParentFolderIds: [rootReference()],
    Traversal: "Deep",
    Paging: {
      __type: "IndexedPageView:#Exchange", BasePoint: "Beginning",
      Offset: offset, MaxEntriesReturned: 100,
    },
  });
}

export function listRequest(folderId, count, offset) {
  parseCount(count);
  assertOffset(offset);
  return envelope("FindItem", {
    ItemShape: {
      __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly",
      AdditionalProperties: metadata.map(property),
    },
    ParentFolderIds: [folderReference(folderId)],
    Traversal: "Shallow",
    Paging: {
      __type: "IndexedPageView:#Exchange", BasePoint: "Beginning",
      Offset: offset, MaxEntriesReturned: Math.min(count, 100),
    },
    SortOrder: [{
      __type: "SortResults:#Exchange",
      Path: property("item:DateTimeReceived"), Order: "Descending",
    }],
  });
}

export function bodyRequest(ids, allowedIds) {
  if (!(allowedIds instanceof Set) || !Array.isArray(ids) || ids.length === 0 || ids.length > 100 ||
      ids.some(id => typeof id !== "string" || !allowedIds.has(id)) ||
      new Set(ids).size !== ids.length) {
    throw new ApiError("OUTSIDE_FOLDER", "Only IDs from this request's verified folder listing may be read.", 403);
  }
  return envelope("GetItem", {
    ItemShape: {
      __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly", BodyType: "Text",
      AdditionalProperties: [...metadata, "item:Body", "message:ToRecipients", "message:CcRecipients"].map(property),
    },
    ItemIds: ids.map(Id => ({ __type: "ItemId:#Exchange", Id })),
  });
}

export function assertReadRequest(action, payload, context = {}) {
  let expected;
  switch (action) {
    case "GetFolder":
      expected = folderRequest(context.folderId);
      break;
    case "FindFolder":
      expected = findFoldersRequest(payload?.Body?.Paging?.Offset);
      break;
    case "FindItem": {
      const paging = payload?.Body?.Paging;
      if (!Number.isSafeInteger(paging?.MaxEntriesReturned) ||
          paging.MaxEntriesReturned < 1 || paging.MaxEntriesReturned > 100) {
        throw new ApiError("REQUEST_REFUSED", "Invalid internal page size.", 403);
      }
      expected = listRequest(context.folderId, paging.MaxEntriesReturned, paging.Offset);
      break;
    }
    case "GetItem": {
      assertFolderId(context.folderId);
      const ids = payload?.Body?.ItemIds;
      expected = bodyRequest(Array.isArray(ids) ? ids.map(item => item?.Id) : undefined, context.allowedIds);
      break;
    }
    default:
      throw new ApiError("ACTION_REFUSED", "Only the allowlisted Outlook read actions are allowed.", 403);
  }
  if (!isDeepStrictEqual(expected, payload)) {
    throw new ApiError("REQUEST_REFUSED", "A request tried to override its folder-scoped read-only template.", 403);
  }
}

export function responseMessages(data, action) {
  const items = data?.Body?.ResponseMessages?.Items;
  if (!Array.isArray(items) || items.length === 0) {
    throw new ApiError("INVALID_OUTLOOK_RESPONSE", `${action} returned an unrecognized response.`);
  }
  for (const item of items) {
    if (item.ResponseClass !== "Success" || item.ResponseCode !== "NoError") {
      const code = typeof item.ResponseCode === "string" && /^[A-Za-z0-9_]+$/.test(item.ResponseCode)
        ? item.ResponseCode : "UnknownError";
      if (["GetFolder", "FindItem"].includes(action) &&
          ["ErrorFolderNotFound", "ErrorParentFolderNotFound"].includes(code)) {
        throw new ApiError("FOLDER_NOT_FOUND", "Folder doesn't exist.", 404);
      }
      throw new ApiError("OUTLOOK_ERROR", `${action} failed: ${code}.`);
    }
  }
  return items;
}

export function assertFolderItem(item, folderId) {
  if (typeof item?.ItemId?.Id !== "string" || item.ParentFolderId?.Id !== folderId) {
    throw new ApiError("FOLDER_MISMATCH", "Outlook returned an item outside the requested folder; no emails returned.", 409);
  }
  if (typeof item.DateTimeReceived !== "string" || !Number.isFinite(Date.parse(item.DateTimeReceived)) ||
      typeof item.IsRead !== "boolean" || typeof item.HasAttachments !== "boolean") {
    throw new ApiError("INVALID_OUTLOOK_RESPONSE", "An email is missing required metadata.");
  }
}

const mailbox = value => value?.Mailbox ?? value;
const address = value => {
  const entry = mailbox(value);
  if (!entry || typeof entry !== "object") return null;
  return { name: entry.Name ?? null, email: entry.EmailAddress ?? null };
};
const addresses = values => {
  if (values === undefined) return [];
  if (!Array.isArray(values)) {
    throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook returned invalid recipients.");
  }
  return values.map(address);
};

function project(item, includeBody) {
  const result = {
    id: item.ItemId.Id,
    subject: item.Subject ?? "",
    from: address(item.From ?? item.Sender),
    receivedAt: item.DateTimeReceived,
    preview: item.Preview ?? null,
    isRead: item.IsRead,
    hasAttachments: item.HasAttachments,
  };
  if (includeBody) {
    if (typeof item.Body?.Value !== "string" ||
        !["Text", "HTML"].includes(item.Body.BodyType)) {
      throw new ApiError("MISSING_BODY", "Outlook did not return a complete message body.");
    }
    result.body = { contentType: item.Body.BodyType, content: item.Body.Value };
    result.to = addresses(item.ToRecipients);
    result.cc = addresses(item.CcRecipients);
  }
  return result;
}

export class FolderReader {
  constructor(sendRead) {
    this.sendRead = sendRead;
  }

  async resolveFolder(rawFolder) {
    const name = parseFolderName(rawFolder);
    const matches = [];
    const seen = new Set();
    let offset = 0;
    for (;;) {
      const responses = responseMessages(
        await this.sendRead("FindFolder", findFoldersRequest(offset)), "FindFolder");
      if (responses.length !== 1) {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "FindFolder returned multiple folder pages.");
      }
      const root = responses[0].RootFolder;
      if (!Array.isArray(root?.Folders) || root.Folders.length > 100 ||
          typeof root.IncludesLastItemInRange !== "boolean") {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook returned an invalid folder listing.");
      }
      for (const folder of root.Folders) {
        if (typeof folder?.FolderId?.Id !== "string" || !folder.FolderId.Id ||
            typeof folder.DisplayName !== "string") {
          throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook returned invalid folder metadata.");
        }
        if (seen.has(folder.FolderId.Id)) {
          throw new ApiError("FOLDER_CHANGED", "The folder listing changed while paging; retry the read.", 409);
        }
        seen.add(folder.FolderId.Id);
        if (nameKey(folder.DisplayName) === nameKey(name)) matches.push(folder);
      }
      if (root.IncludesLastItemInRange) break;
      if (root.Folders.length === 0 || !Number.isSafeInteger(root.IndexedPagingOffset) ||
          root.IndexedPagingOffset <= offset) {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook folder paging did not advance.");
      }
      offset = root.IndexedPagingOffset;
    }
    if (matches.length === 0) {
      throw new ApiError("FOLDER_NOT_FOUND", "Folder doesn't exist.", 404);
    }
    if (matches.length !== 1) {
      throw new ApiError("AMBIGUOUS_FOLDER", "More than one folder has that name. Use a unique folder name.", 409);
    }
    const folderId = matches[0].FolderId.Id;
    const folders = responseMessages(
      await this.sendRead("GetFolder", folderRequest(folderId), { folderId }), "GetFolder")
      .flatMap(item => item.Folders ?? []);
    if (folders.length !== 1 || folders[0].FolderId?.Id !== folderId ||
        typeof folders[0].DisplayName !== "string" || nameKey(folders[0].DisplayName) !== nameKey(name)) {
      throw new ApiError("FOLDER_CHANGED", "The requested folder changed during lookup; retry the read.", 409);
    }
    if (typeof folders[0].FolderClass !== "string" ||
        !/^IPF\.Note(?:\.|$)/i.test(folders[0].FolderClass)) {
      throw new ApiError("NOT_MAIL_FOLDER", "The requested folder is not an email folder.", 400);
    }
    return { id: folderId, name: folders[0].DisplayName };
  }

  async emails(rawFolder, rawCount, includeBody = true) {
    const name = parseFolderName(rawFolder);
    const count = parseCount(rawCount);
    if (typeof includeBody !== "boolean") {
      throw new ApiError("INVALID_INCLUDE_BODY", "includeBody must be true or false.", 400);
    }
    const folder = await this.resolveFolder(name);
    const selected = [];
    const allowedIds = new Set();
    const context = { folderId: folder.id, allowedIds };
    let offset = 0;
    let hasMore = false;
    let previousTime = Infinity;
    while (selected.length < count) {
      const limit = Math.min(count - selected.length, 100);
      const responses = responseMessages(
        await this.sendRead("FindItem", listRequest(folder.id, limit, offset), context), "FindItem");
      if (responses.length !== 1) {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "FindItem returned multiple folder results.");
      }
      const root = responses[0].RootFolder;
      if (!Array.isArray(root?.Items) || typeof root.IncludesLastItemInRange !== "boolean" ||
          root.Items.length > limit) {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook returned an invalid folder page.");
      }
      for (const item of root.Items) {
        assertFolderItem(item, folder.id);
        const time = Date.parse(item.DateTimeReceived);
        if (allowedIds.has(item.ItemId.Id) || time > previousTime) {
          throw new ApiError("FOLDER_CHANGED", "The folder changed while paging; retry the read.", 409);
        }
        previousTime = time;
        allowedIds.add(item.ItemId.Id);
        selected.push(item);
      }
      hasMore = !root.IncludesLastItemInRange;
      if (!hasMore || selected.length === count) break;
      if (root.Items.length === 0 || !Number.isSafeInteger(root.IndexedPagingOffset) ||
          root.IndexedPagingOffset <= offset) {
        throw new ApiError("INVALID_OUTLOOK_RESPONSE", "Outlook paging did not advance.");
      }
      offset = root.IndexedPagingOffset;
    }

    let items = selected;
    if (includeBody) {
      const full = new Map();
      const selectedById = new Map(selected.map(item => [item.ItemId.Id, item]));
      for (let start = 0; start < selected.length; start += 100) {
        const ids = selected.slice(start, start + 100).map(item => item.ItemId.Id);
        const expected = new Set(ids);
        const responses = responseMessages(
          await this.sendRead("GetItem", bodyRequest(ids, allowedIds), context), "GetItem");
        for (const item of responses.flatMap(response => response.Items ?? [])) {
          assertFolderItem(item, folder.id);
          const id = item.ItemId.Id;
          if (!expected.has(id) || full.has(id)) {
            throw new ApiError("FOLDER_MISMATCH", "Outlook returned an unexpected message; no emails returned.", 409);
          }
          const listed = selectedById.get(id);
          if (item.IsRead !== listed.IsRead || item.DateTimeReceived !== listed.DateTimeReceived) {
            throw new ApiError("MESSAGE_CHANGED", "An email's metadata changed during the request.", 409);
          }
          full.set(id, item);
        }
        if (ids.some(id => !full.has(id))) {
          throw new ApiError("MISSING_EMAIL", "An email disappeared while reading; no partial result returned.", 409);
        }
      }
      items = selected.map(item => full.get(item.ItemId.Id));
    }
    return {
      folder: folder.name, requestedCount: count, returnedCount: items.length,
      newestFirst: true, hasMore, emails: items.map(item => project(item, includeBody)),
    };
  }
}
