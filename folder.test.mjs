import test from "node:test";
import assert from "node:assert/strict";
import {
  ApiError, FolderReader,
  assertReadRequest, bodyRequest, findFoldersRequest, folderRequest, listRequest,
  parseCount, parseFolderName, parseIncludeBody,
} from "./folder.mjs";

const FOLDER_ID = "test-folder-hammer";
const FOLDER_NAME = "HAMMER-HAI";
const makeFolder = (name = FOLDER_NAME, id = FOLDER_ID) => ({
  FolderId: { Id: id }, DisplayName: name, FolderClass: "IPF.Note",
});
const response = properties => ({
  Body: { ResponseMessages: { Items: [{ ResponseClass: "Success", ResponseCode: "NoError", ...properties }] } },
});
const email = (index, folderId = FOLDER_ID) => ({
  ItemId: { Id: `${folderId === FOLDER_ID ? "" : `${folderId}-`}item-${index}` },
  ParentFolderId: { Id: folderId },
  Subject: `Subject ${index}`, DateTimeReceived: new Date(Date.UTC(2026, 9, 1) - index * 60000).toISOString(),
  From: { Mailbox: { Name: "Example Sender", EmailAddress: "sender@example.test" } },
  Preview: "Example preview", IsRead: index % 2 === 0, HasAttachments: false,
  Body: { BodyType: "Text", Value: `Full body ${index}` },
});
function mailbox(size, alter = () => {}, folders = [makeFolder()]) {
  const itemsByFolder = new Map(folders.map(folder => [
    folder.FolderId.Id,
    Array.from({ length: size }, (_, index) => email(index, folder.FolderId.Id)),
  ]));
  const calls = [];
  const send = async (action, payload, context = {}) => {
    assertReadRequest(action, payload, context);
    calls.push({ action, payload, folderId: context.folderId });
    let result;
    if (action === "FindFolder") {
      const { Offset, MaxEntriesReturned } = payload.Body.Paging;
      const end = Offset + MaxEntriesReturned;
      result = response({ RootFolder: {
        Folders: folders.slice(Offset, end), IndexedPagingOffset: end,
        IncludesLastItemInRange: end >= folders.length,
      } });
    } else if (action === "GetFolder") {
      result = response({ Folders: folders.filter(folder => folder.FolderId.Id === context.folderId) });
    } else if (action === "FindItem") {
      const items = itemsByFolder.get(context.folderId);
      const { Offset, MaxEntriesReturned } = payload.Body.Paging;
      const end = Offset + MaxEntriesReturned;
      result = response({
        RootFolder: {
          Items: items.slice(Offset, end), IndexedPagingOffset: end,
          IncludesLastItemInRange: end >= items.length,
        },
      });
    } else if (action === "GetItem") {
      const items = itemsByFolder.get(context.folderId);
      result = response({ Items: payload.Body.ItemIds.map(({ Id }) => items.find(item => item.ItemId.Id === Id)) });
    } else {
      assert.fail(`Unexpected action ${action}`);
    }
    result = structuredClone(result);
    alter(action, result, calls.length, payload);
    return result;
  };
  return { reader: new FolderReader(send), calls };
}
const rejectsCode = (fn, code) => assert.rejects(fn, error => error instanceof ApiError && error.code === code);

test("count accepts positive safe integers only", () => {
  for (const value of [1, "3", Number.MAX_SAFE_INTEGER]) assert.equal(parseCount(value), Number(value));
  for (const value of [undefined, null, 0, -1, 1.5, "1.5", "1e2", " 3", "+3", "03", Infinity, NaN, true, [], {}, "9007199254740992"]) {
    assert.throws(() => parseCount(value), error => error.code === "INVALID_COUNT");
  }
});

test("invalid count is rejected before any Outlook request", async () => {
  const { reader, calls } = mailbox(3);
  await rejectsCode(() => reader.emails(FOLDER_NAME, 0), "INVALID_COUNT");
  assert.equal(calls.length, 0);
});

test("list returns exactly the requested newest individual emails without bodies", async () => {
  const { reader, calls } = mailbox(7);
  const result = await reader.emails(FOLDER_NAME, "3", false);
  assert.equal(result.folder, FOLDER_NAME);
  assert.equal(result.requestedCount, 3);
  assert.equal(result.returnedCount, 3);
  assert.equal(result.hasMore, true);
  assert.deepEqual(result.emails.map(item => item.id), ["item-0", "item-1", "item-2"]);
  assert.ok(result.emails.every(item => !("body" in item)));
  assert.deepEqual(calls.map(call => call.action), ["FindFolder", "GetFolder", "FindItem"]);
  assert.equal(calls[2].payload.Body.ParentFolderIds[0].Id, FOLDER_ID);
});

test("read fetches complete bodies only for listed IDs and preserves read status", async () => {
  const { reader, calls } = mailbox(7);
  const result = await reader.emails(FOLDER_NAME, 3, true);
  assert.deepEqual(result.emails.map(item => item.body.content), ["Full body 0", "Full body 1", "Full body 2"]);
  assert.deepEqual(result.emails.map(item => item.isRead), [true, false, true]);
  assert.deepEqual(calls.map(call => call.action), ["FindFolder", "GetFolder", "FindItem", "GetItem"]);
  assert.deepEqual(calls[3].payload.Body.ItemIds.map(item => item.Id), ["item-0", "item-1", "item-2"]);
});

test("fewer available emails returns the actual count and no false hasMore", async () => {
  const result = await mailbox(2).reader.emails(FOLDER_NAME, 10, true);
  assert.equal(result.requestedCount, 10);
  assert.equal(result.returnedCount, 2);
  assert.equal(result.hasMore, false);
});

test("empty folder does not trigger a body request", async () => {
  const { reader, calls } = mailbox(0);
  const result = await reader.emails(FOLDER_NAME, 3, true);
  assert.equal(result.returnedCount, 0);
  assert.deepEqual(result.emails, []);
  assert.equal(result.hasMore, false);
  assert.ok(calls.every(call => call.action !== "GetItem"));
});

test("pagination and body batches respect the count across more than 100 items", async () => {
  const { reader, calls } = mailbox(150);
  const result = await reader.emails(FOLDER_NAME, 103, true);
  assert.equal(result.returnedCount, 103);
  assert.equal(result.hasMore, true);
  assert.deepEqual(calls.filter(call => call.action === "FindItem").map(call => call.payload.Body.Paging), [
    { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: 100 },
    { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 100, MaxEntriesReturned: 3 },
  ]);
  assert.deepEqual(calls.filter(call => call.action === "GetItem").map(call => call.payload.Body.ItemIds.length), [100, 3]);
});

test("renamed or mismatched folder fails before listing", async () => {
  const { reader, calls } = mailbox(3, (action, result) => {
    if (action === "GetFolder") result.Body.ResponseMessages.Items[0].Folders[0].DisplayName = "Other folder";
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3), "FOLDER_CHANGED");
  assert.equal(calls.length, 2);
});

test("an out-of-folder list item cannot be used to fetch bodies", async () => {
  const { reader, calls } = mailbox(3, (action, result) => {
    if (action === "FindItem") result.Body.ResponseMessages.Items[0].RootFolder.Items[0].ParentFolderId.Id = "other";
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3, true), "FOLDER_MISMATCH");
  assert.ok(calls.every(call => call.action !== "GetItem"));
});

test("a message moved during body retrieval fails without returning a partial result", async () => {
  const { reader } = mailbox(3, (action, result) => {
    if (action === "GetItem") result.Body.ResponseMessages.Items[0].Items[0].ParentFolderId.Id = "other";
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3, true), "FOLDER_MISMATCH");
});

test("an unexpected body ID is rejected", async () => {
  const { reader } = mailbox(3, (action, result) => {
    if (action === "GetItem") result.Body.ResponseMessages.Items[0].Items[0].ItemId.Id = "other";
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3, true), "FOLDER_MISMATCH");
});

test("missing body or missing email is an explicit error, not a preview fallback", async () => {
  const missingBody = mailbox(3, (action, result) => {
    if (action === "GetItem") delete result.Body.ResponseMessages.Items[0].Items[0].Body;
  });
  await rejectsCode(() => missingBody.reader.emails(FOLDER_NAME, 3, true), "MISSING_BODY");
  const missingItem = mailbox(3, (action, result) => {
    if (action === "GetItem") result.Body.ResponseMessages.Items[0].Items.pop();
  });
  await rejectsCode(() => missingItem.reader.emails(FOLDER_NAME, 3, true), "MISSING_EMAIL");
});

test("changed read status during retrieval is reported explicitly", async () => {
  const { reader } = mailbox(3, (action, result) => {
    if (action === "GetItem") result.Body.ResponseMessages.Items[0].Items[0].IsRead = false;
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3, true), "MESSAGE_CHANGED");
});

test("duplicate or non-newest-first results fail closed", async () => {
  const { reader } = mailbox(3, (action, result) => {
    if (action === "FindItem") result.Body.ResponseMessages.Items[0].RootFolder.Items.reverse();
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3), "FOLDER_CHANGED");
  const duplicate = mailbox(3, (action, result) => {
    if (action === "FindItem") {
      const items = result.Body.ResponseMessages.Items[0].RootFolder.Items;
      items[1] = structuredClone(items[0]);
    }
  });
  await rejectsCode(() => duplicate.reader.emails(FOLDER_NAME, 3), "FOLDER_CHANGED");
});

test("malformed, oversized, or non-advancing pages are rejected", async () => {
  const oversized = mailbox(3, (action, result) => {
    if (action === "FindItem") result.Body.ResponseMessages.Items[0].RootFolder.Items.push(email(10));
  });
  await rejectsCode(() => oversized.reader.emails(FOLDER_NAME, 3), "INVALID_OUTLOOK_RESPONSE");
  const stalled = mailbox(101, (action, result) => {
    if (action === "FindItem") result.Body.ResponseMessages.Items[0].RootFolder.IndexedPagingOffset = 0;
  });
  await rejectsCode(() => stalled.reader.emails(FOLDER_NAME, 101), "INVALID_OUTLOOK_RESPONSE");
  const malformed = new FolderReader(async () => ({}));
  await rejectsCode(() => malformed.emails(FOLDER_NAME, 3), "INVALID_OUTLOOK_RESPONSE");
});

test("Outlook errors are never reported as an empty successful result", async () => {
  const reader = new FolderReader(async () => ({
    Body: { ResponseMessages: { Items: [{ ResponseClass: "Error", ResponseCode: "ErrorAccessDenied" }] } },
  }));
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3), "OUTLOOK_ERROR");
});

test("all write actions and changes to fixed request templates are refused", () => {
  for (const action of ["SendItem", "UpdateItem", "DeleteItem", "MoveItem", "MarkAsRead", "CreateItem"]) {
    assert.throws(() => assertReadRequest(action, folderRequest()), error => error.code === "ACTION_REFUSED");
  }
  const edits = [
    request => { request.Body.ParentFolderIds[0].Id = "another-folder"; },
    request => { request.Body.Traversal = "Deep"; },
    request => { request.Body.MarkAsRead = true; },
    request => { request.Header.Other = "override"; },
    request => { request.Body.ItemShape.AdditionalProperties.push({ FieldURI: "item:Body" }); },
    request => { request.Body.SortOrder[0].Order = "Ascending"; },
    request => { request.Body.Paging.MaxEntriesReturned = 101; },
  ];
  for (const edit of edits) {
    const request = listRequest(FOLDER_ID, 3, 0);
    edit(request);
    assert.throws(() => assertReadRequest("FindItem", request, { folderId: FOLDER_ID }));
  }
});

test("body reads cannot accept arbitrary or duplicated item IDs", () => {
  assert.throws(() => bodyRequest(["other"], new Set(["item-0"])), error => error.code === "OUTSIDE_FOLDER");
  assert.throws(() => bodyRequest(["item-0", "item-0"], new Set(["item-0"])));
  const request = bodyRequest(["item-0"], new Set(["item-0"]));
  assert.throws(() => assertReadRequest("GetItem", request, { folderId: FOLDER_ID }), error => error.code === "OUTSIDE_FOLDER");
  const context = { folderId: FOLDER_ID, allowedIds: new Set(["item-0"]) };
  assert.doesNotThrow(() => assertReadRequest("GetItem", request, context));
  request.Body.MarkAsRead = true;
  assert.throws(() => assertReadRequest("GetItem", request, context));
});

test("folder names and includeBody are validated before any Outlook request", async () => {
  assert.equal(parseFolderName("  HAMMER-HAI  "), FOLDER_NAME);
  for (const value of [undefined, null, "", "   ", "Inbox\n", 12]) {
    assert.throws(() => parseFolderName(value), error => error.code === "INVALID_FOLDER");
  }
  assert.equal(parseIncludeBody(undefined), true);
  assert.equal(parseIncludeBody("true"), true);
  assert.equal(parseIncludeBody("false"), false);
  for (const value of ["", "yes", "True", "0", "1"]) {
    assert.throws(() => parseIncludeBody(value), error => error.code === "INVALID_INCLUDE_BODY");
  }
  const { reader, calls } = mailbox(3);
  await rejectsCode(() => reader.emails("", 3), "INVALID_FOLDER");
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3, "false"), "INVALID_INCLUDE_BODY");
  assert.equal(calls.length, 0);
});

test("missing folder returns the exact 404 error and never fetches emails", async () => {
  const { reader, calls } = mailbox(3);
  await assert.rejects(() => reader.emails("Missing folder", 3), error =>
    error.code === "FOLDER_NOT_FOUND" && error.status === 404 && error.message === "Folder doesn't exist.");
  assert.deepEqual(calls.map(call => call.action), ["FindFolder"]);
});

test("request names are case-insensitive and resolve to canonical display names", async () => {
  const result = await mailbox(3).reader.emails(" hammer-hai ", 1, false);
  assert.equal(result.folder, FOLDER_NAME);
});

test("different requests are scoped to their own selected folder and item IDs", async () => {
  const { reader, calls } = mailbox(3, () => {}, [makeFolder(), makeFolder("Project A", "project-folder")]);
  const first = await reader.emails(FOLDER_NAME, 1, true);
  const second = await reader.emails("Project A", 2, true);
  assert.equal(first.folder, FOLDER_NAME);
  assert.equal(second.folder, "Project A");
  assert.equal(second.returnedCount, 2);
  assert.deepEqual(second.emails.map(item => item.id), ["project-folder-item-0", "project-folder-item-1"]);
  assert.deepEqual(calls.filter(call => call.action === "FindItem")
    .map(call => call.payload.Body.ParentFolderIds[0].Id), [FOLDER_ID, "project-folder"]);
});

test("duplicate display names fail explicitly without reading either folder", async () => {
  const { reader, calls } = mailbox(3, () => {}, [makeFolder(), makeFolder("hammer-hai", "duplicate")]);
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3), "AMBIGUOUS_FOLDER");
  assert.ok(calls.every(call => call.action === "FindFolder"));
});

test("folder discovery pages through metadata and finds nested or later folders", async () => {
  const folders = Array.from({ length: 101 }, (_, index) => makeFolder(`Project ${index}`, `folder-${index}`));
  folders.push(makeFolder());
  const { reader, calls } = mailbox(1, () => {}, folders);
  assert.equal((await reader.emails(FOLDER_NAME, 1)).returnedCount, 1);
  assert.deepEqual(calls.filter(call => call.action === "FindFolder").map(call => call.payload.Body.Paging.Offset), [0, 100]);
  assert.ok(calls.filter(call => call.action === "FindFolder").every(call => call.payload.Body.Traversal === "Deep"));
});

test("non-email folders fail explicitly before listing items", async () => {
  const folder = { ...makeFolder("Calendar"), FolderClass: "IPF.Appointment" };
  const { reader, calls } = mailbox(3, () => {}, [folder]);
  await rejectsCode(() => reader.emails("Calendar", 3), "NOT_MAIL_FOLDER");
  assert.ok(calls.every(call => !["FindItem", "GetItem"].includes(call.action)));
});

test("folder discovery rejects malformed, repeated, and non-advancing pages", async () => {
  const folders = Array.from({ length: 101 }, (_, index) => makeFolder(`Project ${index}`, `folder-${index}`));
  const stalled = mailbox(1, (action, result) => {
    if (action === "FindFolder") result.Body.ResponseMessages.Items[0].RootFolder.IndexedPagingOffset = 0;
  }, folders);
  await rejectsCode(() => stalled.reader.emails("Project 100", 1), "INVALID_OUTLOOK_RESPONSE");
  const duplicate = mailbox(1, (action, result) => {
    if (action === "FindFolder") result.Body.ResponseMessages.Items[0].RootFolder.Folders.push(makeFolder());
  });
  await rejectsCode(() => duplicate.reader.emails(FOLDER_NAME, 1), "FOLDER_CHANGED");
});

test("folder discovery and item templates cannot be repurposed for another root or folder", () => {
  const discovery = findFoldersRequest(0);
  discovery.Body.ParentFolderIds[0].Id = "another-root";
  assert.throws(() => assertReadRequest("FindFolder", discovery), error => error.code === "REQUEST_REFUSED");
  const listing = listRequest("folder-one", 1, 0);
  assert.throws(() => assertReadRequest("FindItem", listing, { folderId: "folder-two" }),
    error => error.code === "REQUEST_REFUSED");
  const lookup = folderRequest("folder-one");
  assert.throws(() => assertReadRequest("GetFolder", lookup, { folderId: "folder-two" }),
    error => error.code === "REQUEST_REFUSED");
});

test("a folder deleted during lookup becomes FOLDER_NOT_FOUND, not an empty success", async () => {
  const { reader } = mailbox(3, (action, result) => {
    if (action === "GetFolder") {
      result.Body.ResponseMessages.Items = [{ ResponseClass: "Error", ResponseCode: "ErrorFolderNotFound" }];
    }
  });
  await rejectsCode(() => reader.emails(FOLDER_NAME, 3), "FOLDER_NOT_FOUND");
});
