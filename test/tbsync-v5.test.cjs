"use strict";

const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const { once } = require("node:events");
const http = require("node:http");
const path = require("node:path");
const { test } = require("node:test");
const { createV5Handler } = require("../tbsync-v5.cjs");

const grant = {
  helperVersion: 2,
  unrestricted: false,
  target: { accountId: "42", accountName: "Test account", exists: true },
  resources: {
    calendars: { folderId: "f-7", folderName: "Test calendar", exists: true, selected: true,
      bound: true, targetType: "calendars" },
  },
};
const calendar = { id: "cal-7", name: "Test calendar", type: "ext-eas4tbsync@jobisoft.de", readOnly: false };
const folder = { accountId: "42", folderId: "f-7", targetID: calendar.id, selected: true,
  targetType: "calendars", status: "ok", lastSyncTime: 1794150000000 };
const event = { id: "uid-1", calendarId: calendar.id, type: "event", format: "ical",
  instance: "2026-11-08T14:00:00Z", item: "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n" };
const range = { start: "2026-11-08T09:00:00-05:00", end: "2026-11-08T10:00:00-05:00" };

const request = (method, params) => ({ jsonrpc: "2.0", id: 1, method, params });
const read = (handle, name, args) => handle(request("tools/call", { name, arguments: args }));
const unpack = reply => JSON.parse(reply.result.content[0].text);
const refusal = (reply, pattern) => {
  assert.equal(reply.result.isError, true);
  assert.match(reply.result.content[0].text, pattern);
};

async function fixture(t, override = {}) {
  const responses = {
    status: grant,
    "calendars.get": calendar,
    "items.query": [event],
    getFolders: { folders: [folder] },
    getChangelog: { owner: "eas", entries: [] },
    ...override,
  };
  const calls = [];
  const server = http.createServer(async (req, res) => {
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    calls.push({ path: req.url, method: req.method, auth: req.headers.authorization, body });
    const response = responses[body.cmd];
    const result = typeof response === "function" ? response() : response;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ ok: true, result }));
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  t.after(() => new Promise(resolve => server.close(resolve)));
  const port = server.address().port;
  return { calls, handle: createV5Handler(port), port };
}

async function runBridge(t, overrides, inputs) {
  const env = { ...process.env, ...overrides };
  if (overrides.TBSYNC_MCP_BACKEND === undefined) delete env.TBSYNC_MCP_BACKEND;
  const child = spawn(process.execPath, [path.join(__dirname, "../mcp-bridge.cjs")], {
    env, stdio: ["pipe", "pipe", "pipe"],
  });
  t.after(() => child.kill());
  let output = "";
  child.stdout.setEncoding("utf8");
  child.stdout.on("data", chunk => { output += chunk; });
  child.stdin.end(inputs.map(input => JSON.stringify(input)).join("\n") + "\n");
  assert.deepEqual(await once(child, "close"), [0, null]);
  return output.trim().split("\n").map(line => JSON.parse(line));
}

test("v5 CLI advertises only reads and completes the handshake offline", async t => {
  const { calls, port } = await fixture(t);
  const blocked = ["createCalendarEvent", "tbsyncModifyEasEventsBySql", "tbsyncDeleteEasEventsNativeBySql",
    "items.create", "items.update", "items.remove", "tbsyncSyncAccount"];
  const inputs = [
    request("initialize", { protocolVersion: "2025-06-18" }),
    { jsonrpc: "2.0", method: "notifications/initialized" },
    { ...request("tools/list"), id: 2 },
    ...blocked.map((name, i) => ({ ...request("tools/call", { name }), id: 3 + i })),
  ];
  const replies = await runBridge(t, { TBSYNC_MCP_BACKEND: "v5", TBSYNC_V5_PORT: String(port) }, inputs);
  assert.equal(replies.length, 2 + blocked.length);
  assert.equal(replies.find(reply => reply.id === 1).result.protocolVersion, "2025-06-18");
  assert.deepEqual(replies.find(reply => reply.id === 2).result.tools.map(tool => tool.name),
    ["tbsyncGetSyncStatus", "listCalendars", "listCalendarEvents"]);
  for (const reply of replies.filter(item => item.id >= 3)) refusal(reply, /reads and sync status only/);
  assert.deepEqual(calls, []);
});

test("the default CLI route still forwards legacy requests to the old extension", async t => {
  const { calls, port } = await fixture(t);
  const input = request("tools/call", { name: "listCalendars" });
  await runBridge(t, { TBSYNC_MCP_PORT: String(port) }, [input]);
  assert.deepEqual(calls, [{ path: "/", method: "POST", auth: undefined, body: input }]);
});

test("reads iCalendar from the granted account, converting offsets to the native UTC range", async t => {
  const { calls, handle } = await fixture(t, {
    "calendars.get": { ...calendar, url: "private://test" },
    "items.query": [event, { ...event, id: "uid-2" }],
  });
  const result = unpack(await read(handle, "listCalendarEvents", { ...range, calendarId: calendar.id, limit: 1 }));
  assert.deepEqual(result, { calendar, ...range, count: 1, truncated: true,
    events: [{ id: event.id, instance: event.instance, ical: event.item }] });
  assert.deepEqual(calls.map(call => call.body.cmd), ["status", "calendars.get", "items.query", "status"]);
  assert.ok(calls.every(call => call.method === "POST" && call.path === "/rpc" && call.auth === "Bearer tbsync"));
  assert.deepEqual(calls[2].body.args, {
    accountId: "42", resource: "calendars", type: "event", expand: true,
    rangeStart: "20261108T140000Z", rangeEnd: "20261108T150000Z",
  });
});

test("sync status exposes only the granted folder and separates queued edits from calendar messages", async t => {
  const { calls, handle } = await fixture(t, {
    getFolders: { folders: [{ ...folder, folderId: "other", status: "private-other-folder" },
      { ...folder, custom: { private: "internal folder data" } }] },
    getChangelog: { owner: "eas", entries: [
      { status: "modified_by_user", itemId: "private-id" },
      { status: "added_by_server" },
      { status: "modified_for_sendMail" },
    ] },
  });
  const result = unpack(await read(handle, "tbsyncGetSyncStatus"));
  assert.deepEqual(result, { accountId: "42", accountName: "Test account", folderId: "f-7",
    folderName: "Test calendar", ready: true, calendar, status: "ok", lastSyncTime: folder.lastSyncTime,
    pendingChanges: 1, pendingCalendarMessages: 1 });
  assert.deepEqual(calls.find(call => call.body.cmd === "getChangelog").body.args, { accountId: "42", folderId: "f-7" });
  assert.deepEqual(calls.find(call => call.body.cmd === "getFolders").body.args, { accountId: "42" });
  assert.ok(calls.every(call => !["syncAccount", "calendars.synchronize"].includes(call.body.cmd)));
});

test("unbound and unrestricted grants never reach calendar commands", async t => {
  const { calls, handle } = await fixture(t, { status: { ...grant, resources: {} } });
  const result = unpack(await read(handle, "tbsyncGetSyncStatus"));
  assert.equal(result.ready, false);
  assert.equal(result.pendingChanges, null);
  refusal(await read(handle, "listCalendars"), /Select a bound EAS calendar/);
  assert.deepEqual(calls.map(call => call.body.cmd), ["status", "status"]);

  const restricted = await fixture(t, { status: { ...grant, unrestricted: true } });
  refusal(await read(restricted.handle, "listCalendars"), /Disable unrestricted mode/);
  assert.deepEqual(restricted.calls.map(call => call.body.cmd), ["status"]);
});

test("rejects scope overrides and ambiguous or excessive ranges before querying any event", async t => {
  const { calls, handle } = await fixture(t);
  for (const args of [
    { ...range, resource: "tasks" }, { ...range, accountId: "other" }, { ...range, limit: 101 },
    { ...range, start: "2026-11-08T09:00:00" }, { ...range, start: "2026-02-30T09:00:00Z" },
    { ...range, end: "2027-03-01T00:00:00Z" },
  ]) {
    refusal(await read(handle, "listCalendarEvents", args), /arguments|limit|range/);
  }
  assert.deepEqual(calls, []);
  refusal(await read(handle, "listCalendarEvents", { ...range, calendarId: "other" }), /outside/);
  assert.deepEqual(calls.map(call => call.body.cmd), ["status", "calendars.get"]);
});

test("discards mismatched results and results whose grant changed in flight", async t => {
  const wrong = await fixture(t, { "items.query": [{ ...event, calendarId: "other" }] });
  refusal(await read(wrong.handle, "listCalendarEvents", range), /outside the expected calendar/);

  let checks = 0;
  const changed = await fixture(t, { status() {
    if (checks++ === 0) return grant;
    return { ...grant, resources: { calendars: { ...grant.resources.calendars, folderId: "different" } } };
  } });
  refusal(await read(changed.handle, "listCalendarEvents", range), /grant changed/);
  assert.deepEqual(changed.calls.map(call => call.body.cmd), ["status", "calendars.get", "items.query", "status"]);
});

test("an unavailable EAS queue is unknown and an oversized helper response fails closed", async t => {
  const unavailable = await fixture(t, { getChangelog: null });
  const status = unpack(await read(unavailable.handle, "tbsyncGetSyncStatus"));
  assert.equal(status.pendingChanges, null);
  assert.equal(status.pendingCalendarMessages, null);

  const oversized = await fixture(t, { "items.query": [{ ...event, item: "x".repeat(2 * 1024 * 1024) }] });
  refusal(await read(oversized.handle, "listCalendarEvents", range), /exceeded 2 MiB/);
});
