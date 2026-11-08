"use strict";

const assert = require("node:assert/strict");
const { test } = require("node:test");

const scope = { accountID: "42", calendarId: "calendar", confirm: true };
const calendar = { id: scope.calendarId, name: "Test calendar", type: "storage",
  enabled: true, readOnly: false };
const grant = {
  helperVersion: 2, unrestricted: false,
  target: { accountId: scope.accountID, accountName: "Test account", exists: true },
  resources: { calendars: { folderId: "f-7", folderName: "Test calendar", exists: true,
    bound: true, selected: true, targetType: "calendars" } },
};
const capabilities = [
  ["getState", null], ["getFolders", null], ["syncAccount", "account"], ["calendars.get", "calendar"],
  ["calendars.synchronize", "calendar"], ["getChangelog", "folder"], ["items.get", "calendar"],
  ["items.create", "calendar"], ["items.update", "calendar"], ["items.remove", "calendar"],
].map(([cmd, scope]) => ({ cmd, scope }));
const createArgs = { ...scope, title: "New test event", allDay: true, date: "2026-11-08", mcpTag: "test-event" };
const editArgs = { ...scope, eventId: "tag:test-event", expectedIcal: "old", title: "Changed" };

async function fixture(options = {}) {
  const { createV5Handler } = await import("../extension/mcp_server/v5.mjs");
  const calls = [];
  const events = new Map(options.events ?? []);
  const folder = { accountId: scope.accountID, folderId: "f-7", targetID: options.targetID ?? calendar.id, selected: true,
    targetType: "calendars", status: "success", lastSyncTime: 100 };
  let queue = options.queue ?? [];
  const result = (id, item) => ({ id, item, calendarId: calendar.id, type: "event", format: "ical" });
  let refreshes = 0;
  let nextFolderStatus;
  const state = { events, folder, calls };
  const rpc = async (cmd, args = {}, timeout) => {
    calls.push({ cmd, args, timeout });
    switch (cmd) {
      case "status":
        return options.grant?.(state) ?? grant;
      case "help":
        return { commands: options.capabilities ?? capabilities };
      case "getState":
        return { accounts: [{ accountId: scope.accountID, provider: "eas" }] };
      case "calendars.get":
        return { ...calendar, readOnly: options.readOnly ?? calendar.readOnly };
      case "getFolders": {
        if (nextFolderStatus) {
          folder.status = nextFolderStatus.shift();
          if (!nextFolderStatus.length) nextFolderStatus = null;
        }
        return { folders: [{ ...folder }] };
      }
      case "getChangelog":
        if (options.queueError) throw new Error("queue unavailable");
        return { owner: "eas", entries: queue };
      case "calendars.synchronize":
        refreshes++;
        if (options.failRefresh === refreshes) {
          nextFolderStatus = ["pending", "error"];
        } else {
          folder.status = "success";
          folder.lastSyncTime++;
          queue = [];
        }
        return null;
      case "syncAccount":
        if (!options.accountNoop) folder.lastSyncTime++;
        return { folders: [{ ...folder }] };
      case "items.get":
        return events.has(args.id) ? result(args.id, events.get(args.id)) : null;
      case "items.create":
      case "items.update":
        events.set(args.id, args.ical);
        queue.push({ status: "modified_by_user", itemId: args.id });
        if (options.queuedMessage) queue.push({ status: "modified_for_sendMail", itemId: args.id });
        if (options.writeError) throw new Error("write RPC timed out");
        return result(args.id, args.ical);
      case "items.remove":
        events.delete(args.id);
        queue.push({ status: "deleted_by_user", itemId: args.id });
        return null;
      default:
        assert.fail(`Unexpected command ${cmd}`);
    }
  };
  const eventCodec = {
    validate() {},
    idForTag: tag => `tag:${tag}`,
    sameSnapshot: (a, b) => a === b,
    matchesWrite: (a, b) => a === b,
    inspect() {},
    create: (_, args) => args.title,
    update: (_, _ical, args) => args.title,
  };
  const handle = createV5Handler(rpc, { eventCodec, wait: options.wait ?? (async () => {}) });
  return { ...state, handle };
}

async function call(runtime, name, args) {
  const response = await runtime.handle({
    jsonrpc: "2.0", id: 1, method: "tools/call", params: { name, arguments: args },
  });
  const text = response.result.content[0].text;
  if (response.result.isError && !text.startsWith("{")) return { isError: true, error: text };
  return JSON.parse(text);
}

const names = runtime => runtime.calls.map(call => call.cmd);

test("returns only the granted account and refuses writes without confirmation or capabilities", async () => {
  const runtime = await fixture();
  assert.deepEqual(await call(runtime, "tbsyncListAccounts"), [{ accountID: "42", accountname: "Test account" }]);
  runtime.calls.length = 0;
  assert.match((await call(runtime, "createCalendarEvent", { ...createArgs, confirm: false })).error, /confirm=true/);
  assert.deepEqual(names(runtime), []);
  assert.match((await call(runtime, "createCalendarEvent", { ...createArgs, accountID: "wrong" })).error, /requested account/);
  assert.deepEqual(names(runtime), ["status"]);

  const missing = await fixture({ capabilities: capabilities.filter(item => item.cmd !== "items.create") });
  assert.match((await call(missing, "createCalendarEvent", createArgs)).error, /scoped capability/);
  assert.deepEqual(names(missing), ["status", "help"]);
});

test("manual account sync requires fresh success for the granted calendar", async () => {
  const good = await fixture();
  assert.equal((await call(good, "tbsyncSyncAccount", scopeForAccount())).lastSyncTime, 101);
  assert.equal(good.calls.find(call => call.cmd === "syncAccount").timeout, 90000);

  const stalled = await fixture({ accountNoop: true });
  assert.match((await call(stalled, "tbsyncSyncAccount", scopeForAccount())).error, /fresh successful sync/);
});

function scopeForAccount() {
  return { accountID: scope.accountID, confirm: scope.confirm };
}

test("create, update and delete all use one scoped refresh before and after the write", async () => {
  for (const [tool, command, args, prior, stored] of [
    ["createCalendarEvent", "items.create", createArgs, [], "New test event"],
    ["updateCalendarEvent", "items.update", editArgs, [["tag:test-event", "old"]], "Changed"],
    ["deleteCalendarEvent", "items.remove", { ...scope, eventId: editArgs.eventId,
      expectedIcal: editArgs.expectedIcal }, [["tag:test-event", "old"]], undefined],
  ]) {
    const runtime = await fixture({ events: prior });
    const result = await call(runtime, tool, args);
    assert.deepEqual(result, { success: true, outcome: "applied", eventId: "tag:test-event",
      preSync: { status: "success", lastSyncTime: 101 },
      postSync: { status: "success", lastSyncTime: 102 }, localReadback: true });
    assert.deepEqual(names(runtime).filter(name => name === command || name === "calendars.synchronize"),
      ["calendars.synchronize", command, "calendars.synchronize"]);
    assert.equal(runtime.events.get("tag:test-event"), stored);
    assert.ok(runtime.calls.filter(call => call.cmd === command || call.cmd === "calendars.synchronize")
      .every(call => call.args.accountId === "42" && call.args.resource === "calendars"));
    assert.ok(!names(runtime).includes("syncAccount"));
  }
});

test("pending work and failed pre-sync leave the calendar untouched", async () => {
  for (const opts of [{ queue: [{ status: "modified_for_sendMail" }] }, { queueError: true }, { failRefresh: 1 }]) {
    const runtime = await fixture(opts);
    const result = await call(runtime, "createCalendarEvent", createArgs);
    assert.match(result.error, /pending edits|queue unavailable|fresh successful sync/);
    assert.ok(!names(runtime).includes("items.create"));
    if (!opts.failRefresh) assert.ok(!names(runtime).includes("calendars.synchronize"));
    assert.equal(runtime.events.size, 0);
  }
});

test("duplicate creates and stale snapshots are refused before a mutation", async () => {
  const runtime = await fixture({ events: [["tag:test-event", "newer"]] });
  const repeated = await call(runtime, "createCalendarEvent", createArgs);
  assert.equal(repeated.outcome, "alreadyExists");
  assert.equal(repeated.success, false);
  assert.match((await call(runtime, "updateCalendarEvent", editArgs)).error, /changed since it was listed/);
  assert.match((await call(runtime, "deleteCalendarEvent", { ...scope, eventId: editArgs.eventId,
    expectedIcal: editArgs.expectedIcal })).error, /changed since it was listed/);
  assert.ok(!names(runtime).some(cmd => ["items.create", "items.update", "items.remove"].includes(cmd)));
});

test("post-sync failure and an uncertain write report the stage without replaying the mutation", async () => {
  for (const [opts, outcome] of [[{ failRefresh: 2 }, "applied"], [{ writeError: true }, "uncertain"]]) {
    const runtime = await fixture(opts);
    const result = await call(runtime, "createCalendarEvent", createArgs);
    assert.equal(result.success, false);
    assert.equal(result.outcome, outcome);
    assert.match(result.error, /Check Outlook before retrying/);
    assert.equal(names(runtime).filter(cmd => cmd === "items.create").length, 1);
    assert.equal(names(runtime).filter(cmd => cmd === "calendars.synchronize").length, 2);
  }
});

test("a changed grant is refused before the write and concurrent writes are serialized", async () => {
  const changed = await fixture({ grant: ({ calls }) => names({ calls }).includes("items.get")
    ? { ...grant, resources: { calendars: { ...grant.resources.calendars, folderId: "another" } } }
    : grant });
  assert.match((await call(changed, "createCalendarEvent", createArgs)).error, /grant changed/);
  assert.ok(!names(changed).includes("items.create"));

  const concurrent = await fixture();
  const first = call(concurrent, "createCalendarEvent", createArgs);
  const second = await call(concurrent, "createCalendarEvent", createArgs);
  assert.match(second.error, /another calendar write or sync/i);
  assert.equal((await first).success, true);
});

test("an unexpected calendar message blocks the explicit post-sync", async () => {
  const runtime = await fixture({ queuedMessage: true });
  const result = await call(runtime, "createCalendarEvent", createArgs);
  assert.equal(result.success, false);
  assert.equal(result.outcome, "applied");
  assert.equal(result.postSync, null);
  assert.match(result.error, /post-sync was withheld/);
  assert.equal(names(runtime).filter(cmd => cmd === "calendars.synchronize").length, 1);
});

const writeCalls = [
  ["createCalendarEvent", createArgs],
  ["updateCalendarEvent", editArgs],
  ["deleteCalendarEvent", { ...scope, eventId: editArgs.eventId, expectedIcal: editArgs.expectedIcal }],
];

test("sync and writes refuse a storage calendar with a mismatched folder target ID", async () => {
  const runtime = await fixture({ targetID: "different-calendar" });
  for (const [name, args] of [["tbsyncSyncAccount", scopeForAccount()], ...writeCalls]) {
    assert.match((await call(runtime, name, args)).error, /granted calendar binding changed/i, name);
  }
  assert.ok(!names(runtime).some(cmd => ["syncAccount", "calendars.synchronize",
    "items.create", "items.update", "items.remove"].includes(cmd)));
});

test("a bound read-only storage calendar permits account sync and refuses all writes", async () => {
  const runtime = await fixture({ readOnly: true });
  assert.equal((await call(runtime, "tbsyncSyncAccount", scopeForAccount())).success, true);
  runtime.calls.length = 0;
  for (const [name, args] of writeCalls) {
    assert.match((await call(runtime, name, args)).error, /read-only/, name);
  }
  assert.ok(!names(runtime).some(cmd => ["syncAccount", "calendars.synchronize",
    "items.create", "items.update", "items.remove"].includes(cmd)));
});
