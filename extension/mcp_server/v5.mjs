import { parseTimestamp } from "./event-ical.mjs";

const PROTOCOL_VERSIONS = [
  "2024-10-07", "2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25",
];
const HELPER_VERSION = 2;
const SYNC_TIMEOUT_MS = 90000;
const CALENDAR_SYNC_TIMEOUT_MS = 60000;
const MAX_RANGE_MS = 31 * 86400000;

const noArgs = { type: "object", properties: {}, additionalProperties: false };
const readTools = [
  {
    name: "tbsyncListAccounts",
    description: "List the one account granted in the TbSync v5 Bridge tab.",
    inputSchema: noArgs,
  },
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
    name: "listEvents",
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

readTools.push({ ...readTools.at(-1), name: "listCalendarEvents" });

const string = { type: "string", minLength: 1 };
const confirmation = {
  accountID: string,
  confirm: { type: "boolean", const: true, description: "Set only after the user confirms this operation." },
};
const target = { ...confirmation, calendarId: string };
const eventFields = {
  title: string,
  description: { type: "string" },
  location: { type: "string" },
  mcpTag: { type: "string", description: "Stable unique tag (3–64 lowercase letters, digits or hyphens)." },
  allDay: { type: "boolean" },
  date: { type: "string", description: "First all-day date (YYYY-MM-DD)." },
  endDate: { type: "string", description: "Exclusive end date for an all-day event." },
  start: { type: "string", description: "Timed event start with an explicit UTC offset or Z." },
  end: { type: "string", description: "Timed event end with an explicit UTC offset or Z." },
};
const existing = {
  eventId: string,
  expectedIcal: { type: "string", minLength: 1, maxLength: 131072,
    description: "The iCalendar snapshot for this event returned by listEvents." },
};
function tool(name, description, properties, required) {
  return { name, description, inputSchema: { type: "object", properties, required, additionalProperties: false } };
}

const mutationCommands = new Map([
  ["createCalendarEvent", "items.create"],
  ["updateCalendarEvent", "items.update"],
  ["deleteCalendarEvent", "items.remove"],
]);
const writeTools = [
  tool("tbsyncSyncAccount", "Sync the granted EAS account and report a fresh or unchanged successful calendar status.",
    confirmation, ["accountID", "confirm"]),
  tool("createCalendarEvent", "Create a tagged standalone event with a confirmed pre-sync and post-sync.",
    { ...target, ...eventFields }, [...Object.keys(target), "title", "mcpTag", "allDay"]),
  tool("updateCalendarEvent", "Patch one standalone event with a confirmed pre-sync, snapshot check and post-sync.",
    { ...target, ...existing, ...eventFields }, [...Object.keys(target), ...Object.keys(existing)]),
  tool("deleteCalendarEvent", "Delete one standalone event with a confirmed pre-sync, snapshot check and post-sync.",
    { ...target, ...existing }, [...Object.keys(target), ...Object.keys(existing)]),
];

const isObject = value => value !== null && typeof value === "object" && !Array.isArray(value);
const nonEmptyString = value => typeof value === "string" && value.length > 0;

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
    accountExists: nonEmptyString(accountId) && status.target.exists === true,
    ready,
  };
}

function findAccount(state, accountId) {
  return Array.isArray(state?.accounts) ? state.accounts.find(row => row?.accountId === accountId) : null;
}

async function getCalendar(rpc, grant) {
  const account = findAccount(await rpc("getState"), grant.accountId);
  if (account?.provider !== "eas") {
    throw new Error("The granted TbSync account is not an EAS account.");
  }

  const platform = await rpc("calendars.get", { accountId: grant.accountId, resource: "calendars" });
  if (!nonEmptyString(platform?.id)) {
    throw new Error("The granted folder does not resolve to a calendar.");
  }
  // Thunderbird can expose a bound EAS calendar as storage. The TbSync
  // account provider and the folder's target ID identify the binding.
  const { id, name, type, readOnly, enabled } = platform;
  const calendar = { id, name, type, readOnly, enabled };
  const folder = await readFolder(rpc, grant, calendar);
  return { calendar, folder };
}

async function assertSameGrant(rpc, grant) {
  const current = await getGrant(rpc);
  if (!current.ready || current.accountId !== grant.accountId || current.folderId !== grant.folderId) {
    throw new Error("The TbSync Bridge grant changed during the request. Check the grant before retrying.");
  }
}

async function readFolder(rpc, grant, calendar) {
  const { folders } = await rpc("getFolders", { accountId: grant.accountId }) ?? {};
  if (!Array.isArray(folders)) throw new Error("TbSync returned an invalid folder snapshot.");
  // The manager RPC includes other rows. Match both account and folder
  // and use only the granted row's public fields.
  const folder = folders.find(row => row?.accountId === grant.accountId && row.folderId === grant.folderId);
  if (folder?.targetID !== calendar.id || folder.selected !== true || folder.targetType !== "calendars") {
    throw new Error("The granted calendar binding changed. Refresh sync status and try again.");
  }
  return folder;
}

async function readQueue(rpc, { accountId, folderId }, eventId) {
  const queue = await rpc("getChangelog", { accountId, folderId });
  if (!Array.isArray(queue?.entries)) throw new Error("The EAS provider queue is unavailable.");
  // Same vocabulary as EAS's changelog. Server tags do not represent writes.
  const changes = queue.entries.filter(row => String(row?.status).endsWith("_by_user"));
  return {
    pendingChanges: changes.length,
    pendingCalendarMessages: queue.entries.filter(row => String(row?.status).endsWith("_for_sendMail")).length,
    ...(eventId ? { unrelatedPending: changes.some(row => row.itemId !== eventId) } : {}),
  };
}

async function getSyncStatus(rpc, grant, calendar, folder) {
  const queue = await readQueue(rpc, grant).catch(() => ({ pendingChanges: null, pendingCalendarMessages: null }));
  const lastSyncTime = Number.isFinite(folder.lastSyncTime) && folder.lastSyncTime > 0
    ? folder.lastSyncTime : null;
  return { ...grant, calendar, status: folder.status ?? null, lastSyncTime, ...queue };
}

async function requireCapabilities(rpc, mutation) {
  const required = new Map([
    ["getState", null], ["getFolders", null], ["calendars.get", "calendar"],
    mutation ? ["calendars.synchronize", "calendar"] : ["syncAccount", "account"],
  ]);
  if (mutation) {
    required.set("getChangelog", "folder");
    required.set("items.get", "calendar");
    required.set(mutation, "calendar");
  }
  const { commands } = await rpc("help") ?? {};
  if (!Array.isArray(commands) || [...required].some(([cmd, scope]) =>
    !commands.some(item => item?.cmd === cmd && item.scope === scope))) {
    throw new Error("The TbSync beta bridge is missing a required scoped capability.");
  }
}

function hasSuccessfulSync(folder) {
  return folder?.selected === true && folder.status === "success" &&
    !folder.error && Number.isFinite(folder.lastSyncTime) && folder.lastSyncTime > 0;
}

function isFreshSync(folder, before) {
  return hasSuccessfulSync(folder) && folder.lastSyncTime > (before.lastSyncTime ?? 0);
}

async function confirmedAccountSync(rpc, grant, calendar) {
  await assertSameGrant(rpc, grant);
  const before = await readFolder(rpc, grant, calendar);
  const result = await rpc("syncAccount", { accountId: grant.accountId }, SYNC_TIMEOUT_MS);
  const completed = Array.isArray(result?.folders)
    ? result.folders.find(folder => folder?.folderId === grant.folderId) : null;
  const failure = "TbSync did not confirm a successful sync for the granted calendar. Check TbSync Manager.";
  if (!hasSuccessfulSync(completed)) throw new Error(failure);

  const fresh = isFreshSync(completed, before);
  let unchanged;
  if (!fresh) {
    if (!hasSuccessfulSync(before) || before.lastSyncTime !== completed.lastSyncTime) throw new Error(failure);
    const queue = await readQueue(rpc, grant);
    if (queue.pendingChanges || queue.pendingCalendarMessages) {
      throw new Error("TbSync reported pending calendar work without a fresh successful sync. Check TbSync Manager.");
    }
    const state = await rpc("getState");
    const account = findAccount(state, grant.accountId);
    if (account?.enabled !== true || account.error) {
      throw new Error("The TbSync account is disabled or has an error. Check TbSync Manager.");
    }
    if (["syncingAccounts", "cancellingAccounts", "maintainingAccounts", "settingUpAccounts", "upgradeAccounts"]
      .some(key => state.transient?.[key]?.includes(grant.accountId))) {
      throw new Error("The TbSync account is busy. Try again after the current operation completes.");
    }
    // The bridge can return without contacting Exchange, including while a
    // legacy migration is pending. Confirm only the observed folder state.
    unchanged = {
      outcome: "unchanged",
      migrationPending: account.legacyImported === true || account.legacyReplayPending === true,
      ...queue,
    };
  }

  await assertSameGrant(rpc, grant);
  const current = await readFolder(rpc, grant, calendar);
  const consistent = fresh ? current.lastSyncTime >= completed.lastSyncTime
    : current.lastSyncTime === completed.lastSyncTime;
  if (!hasSuccessfulSync(current) || !consistent) throw new Error(failure);
  return { status: "success", lastSyncTime: current.lastSyncTime, ...unchanged };
}

async function confirmedCalendarRefresh(rpc, grant, calendar, wait) {
  await assertSameGrant(rpc, grant);
  const before = await readFolder(rpc, grant, calendar);
  // The scoped refresh returns when scheduled. Require the host's folder
  // state to confirm completion before advancing to the next stage.
  await rpc("calendars.synchronize", { accountId: grant.accountId, resource: "calendars" });
  const deadline = Date.now() + CALENDAR_SYNC_TIMEOUT_MS;
  let observedPending = false;
  do {
    const current = await readFolder(rpc, grant, calendar);
    if (isFreshSync(current, before)) {
      await assertSameGrant(rpc, grant);
      return { status: "success", lastSyncTime: current.lastSyncTime };
    }
    observedPending ||= current.status === "pending";
    if ((observedPending && ["error", "aborted", "warning"].includes(current.status)) ||
        (current.status === "warning" && current.lastSyncTime > (before.lastSyncTime ?? 0))) break;
    await wait(750);
  } while (Date.now() < deadline);
  throw new Error("TbSync did not confirm a fresh successful sync for the granted calendar. Check TbSync Manager.");
}

async function readEvent(rpc, scope, calendar, id) {
  const result = await rpc("items.get", { ...scope, id });
  if (result === null) return null;
  if (result?.id !== id || result.calendarId !== calendar.id || result.type !== "event" ||
      result.format !== "ical" || typeof result.item !== "string" || result.instance) {
    throw new Error("TbSync returned an unexpected event or calendar binding.");
  }
  return result.item;
}

async function requireQuietQueue(rpc, grant) {
  const queue = await readQueue(rpc, grant);
  if (queue.pendingChanges || queue.pendingCalendarMessages) {
    throw new Error("The granted calendar has pending edits or calendar messages. Review them before starting a write.");
  }
}

async function runEventWorkflow(rpc, eventCodec, command, args, grant, calendar, timing, wait) {
  const scope = { accountId: grant.accountId, resource: "calendars" };
  const creating = command === "items.create";
  const deleting = command === "items.remove";
  const eventId = creating ? eventCodec.idForTag(args.mcpTag) : args.eventId;
  await requireQuietQueue(rpc, grant);
  const preSync = await confirmedCalendarRefresh(rpc, grant, calendar, wait);
  await requireQuietQueue(rpc, grant);
  const original = await readEvent(rpc, scope, calendar, eventId);
  if (creating && original !== null) {
    return { success: false, outcome: "alreadyExists", eventId, preSync,
      error: "An event with this MCP tag already exists. Read it before retrying." };
  }
  if (!creating && (original === null || !eventCodec.sameSnapshot(original, args.expectedIcal, eventId))) {
    throw new Error("The event is missing or changed since it was listed. Read it again before editing or deleting.");
  }
  const ical = creating ? eventCodec.create(eventId, args, timing)
    : deleting ? undefined : eventCodec.update(eventId, original, args, timing);
  if (deleting) eventCodec.inspect(original, eventId);
  if (!creating && !deleting && eventCodec.sameSnapshot(original, ical, eventId)) {
    return { success: true, outcome: "unchanged", eventId, preSync };
  }

  // Upstream resolves its calendar grant per call. Check the binding again
  // immediately before the write, and validate the reply and readback.
  await assertSameGrant(rpc, grant);
  const { calendar: current, folder: binding } = await getCalendar(rpc, grant);
  if (current.id !== calendar.id || current.readOnly !== false || current.enabled !== true ||
      binding.readOnly === true || binding.downloadOnly === true) {
    throw new Error("The granted calendar changed or stopped accepting writes.");
  }
  await requireQuietQueue(rpc, grant);
  const rechecked = await readEvent(rpc, scope, calendar, eventId);
  if (creating ? rechecked !== null : rechecked === null || !eventCodec.sameSnapshot(rechecked, original, eventId)) {
    throw new Error("The event changed while preparing the write. Read it again before retrying.");
  }
  let mutation = "uncertain";
  let mutationError;
  try {
    const result = await rpc(command, { ...scope, id: eventId, ...(ical ? { ical } : {}) });
    if (!deleting && (result?.id !== eventId || result.calendarId !== calendar.id || result.type !== "event")) {
      throw new Error("The write response did not identify the expected event and calendar.");
    }
    mutation = "applied";
  } catch (error) {
    mutationError = error.message;
  }

  // A rejected or timed-out RPC can have already applied the edit. Attempt a
  // single post-sync with the same grant, and never repeat the mutation.
  let postSync;
  try {
    await assertSameGrant(rpc, grant);
    const queued = await readQueue(rpc, grant, eventId);
    if (queued.pendingCalendarMessages || queued.unrelatedPending) {
      throw new Error("The calendar queued other edits or calendar messages. The explicit post-sync was withheld.");
    }
    postSync = await confirmedCalendarRefresh(rpc, grant, calendar, wait);
    const after = await readQueue(rpc, grant);
    const stored = await readEvent(rpc, scope, calendar, eventId);
    await assertSameGrant(rpc, grant);
    if (after.pendingChanges || after.pendingCalendarMessages ||
        (deleting ? stored !== null : stored === null || !eventCodec.matchesWrite(stored, ical, eventId))) {
      throw new Error("The provider queue or local event readback did not confirm the requested change.");
    }
    if (mutationError) {
      return { success: false, outcome: mutation, eventId, preSync, postSync, localReadback: true,
        error: `${mutationError} Check Outlook before retrying.` };
    }
    return { success: true, outcome: mutation, eventId, preSync, postSync, localReadback: true };
  } catch (error) {
    return { success: false, outcome: mutation, eventId, preSync, postSync: postSync ?? null,
      error: `Post-write verification failed: ${error.message} Check Outlook before retrying.`,
      ...(mutationError ? { mutationError } : {}) };
  }
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

async function listEvents(rpc, calendar, args, query) {
  if (args.calendarId !== undefined && args.calendarId !== calendar.id) {
    throw new Error("The requested calendar is outside the TbSync Bridge grant.");
  }
  // The helper injects its granted calendar ID and forwards other options
  // to Thunderbird's query schema, which rejects accountId.
  const events = await rpc("items.query", { ...query, resource: "calendars" });
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

export function createV5Handler(rpc, { eventCodec, wait } = {}) {
  const tools = eventCodec && wait ? [...readTools, ...writeTools] : readTools;
  let activeMutation = false;

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
    const command = mutationCommands.get(name);
    const syncing = name === "tbsyncSyncAccount";
    if (command || syncing) {
      if (args.confirm !== true || !nonEmptyString(args.accountID)) {
        throw new Error("Provide the granted accountID and explicit confirm=true.");
      }
      if (command && !nonEmptyString(args.calendarId)) throw new Error("calendarId is required.");
      if (command && command !== "items.create" && (!nonEmptyString(args.eventId) ||
          !nonEmptyString(args.expectedIcal) || args.expectedIcal.length > 131072)) {
        throw new Error("eventId and its bounded expectedIcal snapshot are required.");
      }
      const timing = command && command !== "items.remove"
        ? eventCodec.validate(args, command === "items.create") : undefined;
      if (activeMutation) throw new Error("Another calendar write or sync is in progress.");
      activeMutation = true;
      try {
        const grant = await getGrant(rpc);
        if (!grant.ready || grant.accountId !== args.accountID) {
          throw new Error("Select a bound EAS calendar for the requested account in the TbSync Bridge tab.");
        }
        await requireCapabilities(rpc, command);
        const { calendar } = await getCalendar(rpc, grant);
        if (syncing) {
          const sync = await confirmedAccountSync(rpc, grant, calendar);
          return { success: true, accountID: grant.accountId, calendar, ...sync };
        }
        if (calendar.id !== args.calendarId || calendar.readOnly !== false || calendar.enabled !== true) {
          throw new Error("The requested calendar is outside the grant, disabled, or read-only.");
        }
        return await runEventWorkflow(rpc, eventCodec, command, args, grant, calendar, timing, wait);
      } finally {
        activeMutation = false;
      }
    }

    const query = name === "listEvents" || name === "listCalendarEvents" ? eventQuery(args) : null;
    const grant = await getGrant(rpc);
    if (name === "tbsyncListAccounts") {
      return grant.accountExists ? [{ accountID: grant.accountId, accountname: grant.accountName }] : [];
    }
    if (!grant.ready) {
      if (name === "tbsyncGetSyncStatus") {
        return { ...grant, status: null, lastSyncTime: null, pendingChanges: null, pendingCalendarMessages: null };
      }
      throw new Error("Select a bound EAS calendar in the TbSync Bridge tab.");
    }

    const { calendar, folder } = await getCalendar(rpc, grant);
    let result;
    if (name === "tbsyncGetSyncStatus") result = await getSyncStatus(rpc, grant, calendar, folder);
    else if (name === "listCalendars") result = [calendar];
    else result = await listEvents(rpc, calendar, args, query);
    await assertSameGrant(rpc, grant);
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
          serverInfo: { name: "thunderbird-tbsync-mcp", version: "0.11.2" },
        } };
      }
      case "ping":
        return { ...reply, result: {} };
      case "tools/list":
        return { ...reply, result: { tools } };
      case "tools/call":
        try {
          const value = await callTool(params?.name, params?.arguments ?? {});
          return { ...reply, result: {
            ...(value?.success === false ? { isError: true } : {}),
            content: [{ type: "text", text: JSON.stringify(value) }],
          } };
        } catch (error) {
          return { ...reply, result: { isError: true, content: [{ type: "text", text: error.message }] } };
        }
      default:
        return { ...reply, error: { code: -32601, message: "Method not found." } };
    }
  };
}
