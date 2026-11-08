"use strict";

const http = require("node:http");

const PROTOCOL_VERSIONS = [
  "2024-10-07", "2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25",
];
const HELPER_VERSION = 2;
const HELPER_TIMEOUT_MS = 5000;
const MAX_RESPONSE_BYTES = 2 * 1024 * 1024;
const MAX_RANGE_MS = 31 * 86400000;
const EAS_CALENDAR_TYPE = "ext-eas4tbsync@jobisoft.de";

const noArgs = { type: "object", properties: {}, additionalProperties: false };
const tools = [
  {
    name: "tbsyncGetSyncStatus",
    description: "Inspect the granted EAS calendar's binding, last sync and pending changes. Does not trigger sync.",
    inputSchema: noArgs,
  },
  {
    name: "listCalendars",
    description: "List the one EAS calendar granted in the TbSync v5 Bridge tab.",
    inputSchema: noArgs,
  },
  {
    name: "listCalendarEvents",
    description: "Read events and expanded recurrences from the granted EAS calendar as iCalendar. Use an explicit offset or Z and a range up to 31 days.",
    inputSchema: {
      type: "object",
      properties: {
        calendarId: { type: "string", minLength: 1 },
        start: { type: "string", description: "Inclusive start, e.g. 2026-11-08T00:00:00-05:00" },
        end: { type: "string", description: "Exclusive end, e.g. 2026-11-09T00:00:00-05:00" },
        limit: { type: "integer", minimum: 1, maximum: 100, default: 50 },
      },
      required: ["start", "end"],
      additionalProperties: false,
    },
  },
];

const isObject = value => value !== null && typeof value === "object" && !Array.isArray(value);
const nonEmptyString = value => typeof value === "string" && value.length > 0;

function callHelper(port, cmd, args = {}) {
  const payload = Buffer.from(JSON.stringify({ cmd, args, timeoutMs: HELPER_TIMEOUT_MS }));
  return new Promise((resolve, reject) => {
    const request = http.request({
      hostname: "127.0.0.1",
      port,
      path: "/rpc",
      method: "POST",
      // Upstream helper v2 requires this static header to block simple
      // cross-origin browser requests. It is not a local-process secret.
      headers: {
        Authorization: "Bearer tbsync",
        "Content-Type": "application/json",
        "Content-Length": payload.length,
      },
    }, response => {
      const chunks = [];
      let bytes = 0;
      response.on("data", chunk => {
        bytes += chunk.length;
        if (bytes > MAX_RESPONSE_BYTES) {
          request.destroy(new Error("TbSync response exceeded 2 MiB. Try a smaller date range."));
          return;
        }
        chunks.push(chunk);
      });
      response.on("error", reject);
      response.on("end", () => {
        if (response.statusCode !== 200) {
          reject(new Error(`TbSync helper returned HTTP ${response.statusCode} for ${cmd}.`));
          return;
        }
        try {
          const body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
          if (body?.ok !== true || !Object.hasOwn(body, "result")) {
            const code = typeof body?.errorCode === "string" && /^E:[A-Z_]+$/.test(body.errorCode)
              ? ` (${body.errorCode})` : "";
            reject(new Error(`TbSync refused ${cmd}${code}. Check the Bridge tab.`));
            return;
          }
          resolve(body.result);
        } catch {
          reject(new Error("TbSync helper returned invalid JSON."));
        }
      });
    });
    const deadline = setTimeout(() => request.destroy(new Error(`TbSync timed out during ${cmd}.`)),
      HELPER_TIMEOUT_MS + 1000);
    request.once("close", () => clearTimeout(deadline));
    request.on("error", error => {
      reject(error.code === "ECONNREFUSED"
        ? new Error("TbSync v5 helper is unavailable. Check the Bridge tab in the TbSync beta build.")
        : error);
    });
    request.end(payload);
  });
}

async function getGrant(rpc) {
  const status = await rpc("status");
  if (status?.helperVersion !== HELPER_VERSION || !isObject(status.target) || !isObject(status.resources)) {
    throw new Error("Expected the TbSync v5 beta bridge with native helper protocol 2.");
  }
  if (status.unrestricted !== false) {
    throw new Error("Disable unrestricted mode in the TbSync Bridge tab before using this MCP backend.");
  }
  const folder = status.resources.calendars;
  const accountId = status.target.accountId ?? null;
  const folderId = folder?.folderId ?? null;
  const ready = nonEmptyString(accountId) && status.target.exists === true &&
    nonEmptyString(folderId) && folder.exists === true && folder.selected === true &&
    folder.bound === true && folder.targetType === "calendars";
  return {
    accountId,
    accountName: status.target.accountName ?? null,
    folderId,
    folderName: folder?.folderName ?? null,
    ready,
  };
}

async function getCalendar(rpc, grant) {
  const calendar = await rpc("calendars.get", { accountId: grant.accountId, resource: "calendars" });
  if (!nonEmptyString(calendar?.id) || calendar.type !== EAS_CALENDAR_TYPE) {
    throw new Error("The granted folder does not resolve to an EAS calendar.");
  }
  const { id, name, type, readOnly } = calendar;
  return { id, name, type, readOnly };
}

async function getSyncStatus(rpc, grant, calendar) {
  const { accountId, folderId } = grant;
  const [snapshot, queue] = await Promise.all([
    rpc("getFolders", { accountId }),
    rpc("getChangelog", { accountId, folderId }).catch(() => null),
  ]);
  // The manager RPC returns all folder rows for this account. Only inspect
  // the granted row and expose its public sync fields, never its custom data.
  const folders = snapshot?.folders;
  if (!Array.isArray(folders)) throw new Error("TbSync returned an invalid folder snapshot.");
  const folder = folders.find(row => row?.accountId === accountId && row.folderId === folderId);
  if (folder?.targetID !== calendar.id || folder.selected !== true || folder.targetType !== "calendars") {
    throw new Error("The granted calendar binding changed. Refresh sync status and try again.");
  }
  const lastSyncTime = Number.isFinite(folder.lastSyncTime) && folder.lastSyncTime > 0
    ? folder.lastSyncTime : null;
  const entries = Array.isArray(queue?.entries) ? queue.entries : null;
  return {
    ...grant,
    calendar,
    status: folder.status ?? null,
    lastSyncTime,
    // Match EAS's pending queue vocabulary. Server tags and scheduling
    // messages can appear in the raw log alongside actual calendar edits.
    pendingChanges: entries?.filter(row => String(row?.status).endsWith("_by_user")).length ?? null,
    pendingCalendarMessages: entries?.filter(row => String(row?.status).endsWith("_for_sendMail")).length ?? null,
  };
}

function parseTimestamp(value) {
  const pattern = /^\d{4}-\d{2}-\d{2}T(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?(?:Z|[+-](?:[01]\d|2[0-3]):[0-5]\d)$/;
  if (typeof value !== "string" || !pattern.test(value)) return NaN;
  const date = value.slice(0, 10);
  const midnight = new Date(`${date}T00:00:00Z`);
  if (!Number.isFinite(midnight.getTime()) || midnight.toISOString().slice(0, 10) !== date) return NaN;
  const instant = new Date(value);
  const year = instant.getUTCFullYear();
  return year >= 1 && year <= 9999 ? instant.getTime() : NaN;
}

function eventQuery(args) {
  const start = parseTimestamp(args.start);
  const end = parseTimestamp(args.end);
  if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start || end - start > MAX_RANGE_MS) {
    throw new Error("Provide an increasing start and end with a timezone offset or Z and a range up to 31 days.");
  }
  // Thunderbird's cal.createDateTime expects the iCalendar basic format.
  const toICal = ms => new Date(ms).toISOString().replace(/[-:]|\.000/g, "");
  return { type: "event", expand: true, rangeStart: toICal(start), rangeEnd: toICal(end) };
}

async function listEvents(rpc, grant, calendar, args, query) {
  if (args.calendarId !== undefined && args.calendarId !== calendar.id) {
    throw new Error("The requested calendar is outside the TbSync Bridge grant.");
  }
  const events = await rpc("items.query", { ...query, accountId: grant.accountId, resource: "calendars" });
  if (!Array.isArray(events) || events.some(event => event?.calendarId !== calendar.id ||
    event.type !== "event" || event.format !== "ical" || typeof event.item !== "string" ||
    !nonEmptyString(event.id))) {
    throw new Error("TbSync returned events outside the expected calendar or response format.");
  }
  const limit = args.limit ?? 50;
  return {
    calendar,
    start: args.start,
    end: args.end,
    count: Math.min(events.length, limit),
    truncated: events.length > limit,
    events: events.slice(0, limit).map(({ id, instance, item }) => ({ id, instance, ical: item })),
  };
}

function createV5Handler(port = 47654) {
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("TBSYNC_V5_PORT must be a TCP port from 1 to 65535.");
  }
  const rpc = (cmd, args) => callHelper(port, cmd, args);

  async function callTool(name, args) {
    const tool = tools.find(item => item.name === name);
    if (!tool) throw new Error("This TbSync v5 backend supports calendar reads and sync status only.");
    if (!isObject(args) || Object.keys(args).some(key => !Object.hasOwn(tool.inputSchema.properties, key))) {
      throw new Error("Invalid arguments for this read-only tool.");
    }
    if (args.calendarId !== undefined && !nonEmptyString(args.calendarId)) {
      throw new Error("calendarId must be a nonempty string.");
    }
    if (args.limit !== undefined && (!Number.isInteger(args.limit) || args.limit < 1 || args.limit > 100)) {
      throw new Error("limit must be an integer between 1 and 100.");
    }
    const query = name === "listCalendarEvents" ? eventQuery(args) : null;
    const grant = await getGrant(rpc);
    if (!grant.ready) {
      if (name === "tbsyncGetSyncStatus") {
        return { ...grant, status: null, lastSyncTime: null, pendingChanges: null, pendingCalendarMessages: null };
      }
      throw new Error("Select a bound EAS calendar in the TbSync Bridge tab.");
    }

    const calendar = await getCalendar(rpc, grant);
    let result;
    switch (name) {
      case "tbsyncGetSyncStatus":
        result = await getSyncStatus(rpc, grant, calendar);
        break;
      case "listCalendars":
        result = [calendar];
        break;
      case "listCalendarEvents":
        result = await listEvents(rpc, grant, calendar, args, query);
        break;
    }
    const current = await getGrant(rpc);
    if (!current.ready || current.accountId !== grant.accountId || current.folderId !== grant.folderId) {
      throw new Error("The TbSync Bridge grant changed during the request. Try again.");
    }
    return result;
  }

  return async function handle(message) {
    const reply = { jsonrpc: "2.0", id: message?.id ?? null };
    if (!isObject(message) || message.jsonrpc !== "2.0" || typeof message.method !== "string") {
      return { ...reply, error: { code: -32600, message: "Invalid JSON-RPC request." } };
    }
    if (!Object.hasOwn(message, "id")) return null;

    const { method, params } = message;
    switch (method) {
      case "initialize": {
        const requested = params?.protocolVersion;
        if (typeof requested !== "string") {
          return { ...reply, error: { code: -32602, message: "protocolVersion must be a string." } };
        }
        const protocolVersion = PROTOCOL_VERSIONS.includes(requested) ? requested : PROTOCOL_VERSIONS.at(-1);
        return { ...reply, result: {
          protocolVersion,
          capabilities: { tools: {} },
          serverInfo: { name: "thunderbird-tbsync-mcp-v5-readonly", version: "0.1.0" },
        } };
      }
      case "ping":
        return { ...reply, result: {} };
      case "tools/list":
        return { ...reply, result: { tools } };
      case "tools/call":
        try {
          const value = await callTool(params?.name, params?.arguments ?? {});
          return { ...reply, result: { content: [{ type: "text", text: JSON.stringify(value) }] } };
        } catch (error) {
          return { ...reply, result: { isError: true, content: [{ type: "text", text: error.message }] } };
        }
      default:
        return { ...reply, error: { code: -32601, message: "Method not found." } };
    }
  };
}

module.exports = { createV5Handler };
