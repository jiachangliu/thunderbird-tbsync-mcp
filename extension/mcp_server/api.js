/* global ExtensionCommon, ChromeUtils, Services, Cc, Ci */
"use strict";

/**
 * Thunderbird TbSync MCP Server Extension
 *
 * Exposes calendar-sync and calendar-write operations via a small HTTP JSON-RPC
 * server (localhost:8766). Intended to be used behind an MCP stdio bridge.
 *
 * HARD REQUIREMENTS:
 * - TbSync installed (addon id: tbsync@jobisoft.de)
 * - Provider for Exchange ActiveSync installed (addon id: eas4tbsync@jobisoft.de)
 *
 * Design goal:
 * - Provide a safe workflow: sync -> write -> sync
 * - Avoid relying on periodic TbSync sync cycles.
 *
 * Licensing:
 * - This project is MIT licensed.
 * - It bundles Mozilla's httpd.sys.mjs (MPL-2.0) copied from thunderbird-mcp.
 */

const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

const MCP_PORT = 8766;
let _serverStarted = false;
let _serverStarting = false;
let _serverPort = null;

var tbsyncMcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    // Make the bundled httpd.sys.mjs available via resource://thunderbird-tbsync-mcp/...
    const extensionRoot = context.extension.rootURI;
    const resourceName = "thunderbird-tbsync-mcp";

    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    return {
      tbsyncMcpServer: {
        start: async function () {
          try {
            if (_serverStarted) {
              return { success: true, port: _serverPort || MCP_PORT, alreadyStarted: true };
            }
            if (_serverStarting) {
              return { success: true, port: _serverPort || MCP_PORT, starting: true };
            }
            _serverStarting = true;
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-tbsync-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { AddonManager } = ChromeUtils.importESModule(
              "resource://gre/modules/AddonManager.sys.mjs"
            );
            const { ExtensionParent } = ChromeUtils.importESModule(
              "resource://gre/modules/ExtensionParent.sys.mjs"
            );

            let cal = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
            } catch {
              // Calendar not available
            }

            function readRequestBody(request) {
              const stream = request.bodyInputStream;
              return NetUtil.readInputStreamToString(stream, stream.available(), { charset: "UTF-8" });
            }

            function sanitizeForJson(text) {
              if (!text) return text;
              return text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, "");
            }

            // Keep timeout timers strongly referenced; otherwise they can be GC'd before firing.
            const _activeTimers = new Set();
            function timeoutAfter(timeoutMs) {
              return new Promise((resolve) => {
                const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                _activeTimers.add(timer);
                timer.init(
                  {
                    notify: () => {
                      try {
                        _activeTimers.delete(timer);
                      } catch {}
                      resolve({ timeout: true });
                    },
                  },
                  timeoutMs,
                  Ci.nsITimer.TYPE_ONE_SHOT
                );
              });
            }

            // Non-blocking calendar query jobs (avoids hanging MCP calls if providers never complete).
            const _calendarJobs = new Map();
            let _calendarJobSeq = 0;

            // Idempotency cache to prevent duplicate creates when a user repeats the same request.
            // Keyed by calendar+tag+time. TTL-based (in-memory only).
            const _idempotency = new Map();
            const IDEMPOTENCY_TTL_MS = 6 * 60 * 60 * 1000; // 6 hours

            // Workflow jobs (non-blocking sync->create->sync sequences)
            const _workflowJobs = new Map();
            let _workflowSeq = 0;
            function newWorkflowJobId() {
              _workflowSeq += 1;
              return `wfjob-${Date.now()}-${_workflowSeq}`;
            }
            function getWorkflowJob(jobId) {
              const job = _workflowJobs.get(String(jobId));
              if (!job) return { error: `Unknown workflow jobId: ${jobId}` };
              return {
                success: true,
                jobId: job.jobId,
                status: job.status,
                createdAtMs: job.createdAtMs,
                updatedAtMs: job.updatedAtMs,
                step: job.step,
                meta: job.meta,
                result: job.result || null,
                error: job.error || null,
              };
            }

            function containsAll(haystack, needles) {
              const h = String(haystack || "").toLowerCase();
              for (const n of (needles || [])) {
                const nn = String(n || "").toLowerCase().trim();
                if (!nn) continue;
                if (!h.includes(nn)) return false;
              }
              return true;
            }

            async function syncAndDeleteEventsByTitle(args) {
              // Legacy (provider getItems based) bulk delete.
              // NOTE: This can fail for TbSync/EAS calendars because getItems may never return.
              // Prefer syncAndDeleteEventsBySql below.
              const { tbsyncUser, calendarName, calendarId, start, end, titleMustContainAll } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const target = resolveCalendar({ calendarName, calendarId });
              if (!target) return { error: `Calendar not found (name=${calendarName || ""}, id=${calendarId || ""})` };
              if (target.readOnly) return { error: `Calendar is read-only: ${target.name}` };

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              const wfJobId = newWorkflowJobId();
              const job = {
                jobId: wfJobId,
                status: "running",
                createdAtMs: Date.now(),
                updatedAtMs: Date.now(),
                step: "scanAndDelete",
                meta: { tbsyncUser, calendar: { id: target.id, name: target.name }, start, end, titleMustContainAll, mode: "provider" },
                result: { matched: 0, deletedTriggered: 0 },
                error: null,
              };
              _workflowJobs.set(wfJobId, job);

              try {
                maybeRefreshCalendar(target);

                const FILTER = Ci.calICalendar;
                const filter =
                  FILTER.ITEM_FILTER_TYPE_EVENT |
                  FILTER.ITEM_FILTER_CLASS_OCCURRENCES |
                  FILTER.ITEM_FILTER_COMPLETED_YES |
                  FILTER.ITEM_FILTER_COMPLETED_NO;

                const listener = {
                  onOperationComplete: function () {
                    try {
                      job.step = "postSync";
                      job.updatedAtMs = Date.now();
                      try { tbsyncSyncAccountByUser(tbsyncUser); } catch {}
                      try { tbsyncSyncAccountByUser(tbsyncUser); } catch {}
                      job.status = "done";
                      job.step = "done";
                      job.updatedAtMs = Date.now();
                    } catch (e) {
                      job.status = "error";
                      job.step = "error";
                      job.error = e.toString();
                      job.updatedAtMs = Date.now();
                    }
                  },
                  onGetResult: function (_cal, _status, _opType, _id, _detail, count, items) {
                    for (let i = 0; i < count; i++) {
                      const it = items[i];
                      const title = it && it.title ? it.title : "";
                      if (!containsAll(title, titleMustContainAll)) continue;
                      job.result.matched += 1;
                      try {
                        startDeleteItemJob(target, it, { reason: "bulkDelete", title });
                        job.result.deletedTriggered += 1;
                      } catch {}
                    }
                    job.updatedAtMs = Date.now();
                  },
                };

                job._listener = listener;
                target.getItems(filter, 0, range.start, range.end, listener);
              } catch (e) {
                job.status = "error";
                job.step = "error";
                job.updatedAtMs = Date.now();
                job.error = e.toString();
              }

              return { success: true, pending: true, workflowJobId: wfJobId };
            }

            async function syncAndDeleteEventsBySql(args) {
              // Robust bulk delete for TbSync calendars:
              // 1) Query Thunderbird calendar storage DB (calendar-data/local.sqlite) to find candidate item ids
              // 2) Delete those items through the calendar provider (getItem -> deleteItem)
              // 3) Trigger TbSync sync to push deletions to cloud
              const { tbsyncUser, calendarName, calendarId, start, end, titleMustContainAll } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const target = resolveCalendar({ calendarName, calendarId });
              if (!target) return { error: `Calendar not found (name=${calendarName || ""}, id=${calendarId || ""})` };
              if (target.readOnly) return { error: `Calendar is read-only: ${target.name}` };

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              const wfJobId = newWorkflowJobId();
              const job = {
                jobId: wfJobId,
                status: "running",
                createdAtMs: Date.now(),
                updatedAtMs: Date.now(),
                step: "querySql",
                meta: { tbsyncUser, calendar: { id: target.id, name: target.name }, start, end, titleMustContainAll, mode: "sql+providerDelete" },
                result: { candidateCount: 0, deleteTriggered: 0, deleteErrors: 0 },
                error: null,
              };
              _workflowJobs.set(wfJobId, job);

              (async () => {
                try {
                  // Load Sqlite module
                  const { Sqlite } = ChromeUtils.importESModule("resource://gre/modules/Sqlite.sys.mjs");

                  const prof = Services.dirsvc.get("ProfD", Ci.nsIFile);
                  prof.append("calendar-data");
                  prof.append("local.sqlite");
                  const dbPath = prof.path;

                  const conn = await Sqlite.openConnection({ path: dbPath });
                  try {
                    const clauses = [];
                    const params = {
                      cal_id: target.id,
                      start_us: range.start.nativeTime,
                      end_us: range.end.nativeTime,
                    };
                    clauses.push("cal_id = :cal_id");
                    clauses.push("event_start >= :start_us AND event_start < :end_us");

                    // Title contains ALL keywords
                    let idx = 0;
                    for (const kwRaw of (titleMustContainAll || [])) {
                      const kw = String(kwRaw || "").toLowerCase().trim();
                      if (!kw) continue;
                      idx += 1;
                      const key = `kw${idx}`;
                      params[key] = `%${kw}%`;
                      clauses.push(`lower(title) LIKE :${key}`);
                    }

                    const sql = `SELECT id, title FROM cal_events WHERE ${clauses.join(" AND ")}`;
                    const rows = await conn.execute(sql, params);
                    const ids = rows.map((r) => ({ id: r.getResultByName("id"), title: r.getResultByName("title") }));
                    job.result.candidateCount = ids.length;
                    job.updatedAtMs = Date.now();
                    job.step = "delete";

                    // For each id: getItem -> deleteItem
                    job.result.sample = ids.slice(0, 10);
                    job._listeners = [];

                    for (const it of ids) {
                      try {
                        // Many providers only need the id to delete; avoid getItem() which can hang.
                        const dummy = Cc["@mozilla.org/calendar/event;1"].createInstance(Ci.calIEvent);
                        dummy.id = it.id;
                        dummy.title = it.title || "";
                        startDeleteItemJob(target, dummy, { reason: "sqlBulkDelete", title: it.title });
                        job.result.deleteTriggered += 1;
                      } catch (e) {
                        job.result.deleteErrors += 1;
                      }
                      job.updatedAtMs = Date.now();
                    }
                  } finally {
                    await conn.close();
                  }

                  job.step = "postSync";
                  job.updatedAtMs = Date.now();
                  try { tbsyncSyncAccountByUser(tbsyncUser); } catch {}
                  try { tbsyncSyncAccountByUser(tbsyncUser); } catch {}

                  job.status = "done";
                  job.step = "done";
                  job.updatedAtMs = Date.now();
                } catch (e) {
                  job.status = "error";
                  job.step = "error";
                  job.updatedAtMs = Date.now();
                  job.error = e.toString();
                }
              })();

              return { success: true, pending: true, workflowJobId: wfJobId, note: "SQL-based delete started. Poll with getWorkflowJob." };
            }

            function newCalendarJobId() {
              _calendarJobSeq += 1;
              return `caljob-${Date.now()}-${_calendarJobSeq}`;
            }

            function startGetItemsJob(target, filter, startDt, endDt, meta) {
              const jobId = newCalendarJobId();
              const job = {
                jobId,
                status: "pending",
                createdAtMs: Date.now(),
                meta: meta || {},
                events: [],
                error: null,
              };
              _calendarJobs.set(jobId, job);

              try {
                const listener = {
                  onOperationComplete: function () {
                    job.status = "done";
                  },
                  onGetResult: function (_cal, _status, _opType, _id, _detail, count, items) {
                    for (let i = 0; i < count; i++) {
                      const it = items[i];
                      job.events.push({
                        id: it.id || null,
                        title: it.title || null,
                        allDay: it.startDate ? !!it.startDate.isDate : null,
                        start: it.startDate ? it.startDate.icalString : null,
                        end: it.endDate ? it.endDate.icalString : null,
                      });
                    }
                  },
                };
                // Keep listener alive until completion.
                job._listener = listener;
                target.getItems(filter, 0, startDt, endDt, listener);
              } catch (e) {
                job.status = "error";
                job.error = e.toString();
              }

              return job;
            }

            function startAddItemJob(target, item, meta) {
              const jobId = newCalendarJobId();
              const job = {
                jobId,
                status: "pending",
                createdAtMs: Date.now(),
                meta: meta || {},
                events: [],
                error: null,
                itemId: null,
              };
              _calendarJobs.set(jobId, job);

              try {
                const listener = {
                  onOperationComplete: function (_cal, status, _opType, id, _detail) {
                    job.itemId = id || null;
                    // status is an nsresult; if non-zero, treat as error.
                    if (status && status !== 0) {
                      job.status = "error";
                      job.error = `addItem failed with status=${status}`;
                    } else {
                      job.status = "done";
                    }
                  },
                  onGetResult: function () {},
                };
                // Keep listener alive until completion.
                job._listener = listener;
                target.addItem(item, listener);
              } catch (e) {
                job.status = "error";
                job.error = e.toString();
              }

              return job;
            }

            function startDeleteItemJob(target, item, meta) {
              const jobId = newCalendarJobId();
              const job = {
                jobId,
                status: "pending",
                createdAtMs: Date.now(),
                meta: meta || {},
                events: [],
                error: null,
                deletedId: item && item.id ? item.id : null,
              };
              _calendarJobs.set(jobId, job);

              try {
                const listener = {
                  onOperationComplete: function (_cal, status) {
                    if (status && status !== 0) {
                      job.status = "error";
                      job.error = `deleteItem failed with status=${status}`;
                    } else {
                      job.status = "done";
                    }
                  },
                  onGetResult: function () {},
                };
                job._listener = listener;
                target.deleteItem(item, listener);
              } catch (e) {
                job.status = "error";
                job.error = e.toString();
              }

              return job;
            }

            function getCalendarJob(jobId) {
              const job = _calendarJobs.get(String(jobId));
              if (!job) return { error: `Unknown jobId: ${jobId}` };
              return {
                success: true,
                jobId: job.jobId,
                status: job.status,
                createdAtMs: job.createdAtMs,
                meta: job.meta,
                count: job.events.length,
                events: job.events,
                error: job.error,
              };
            }

            async function assertTbSyncInstalled() {
              const tbs = await AddonManager.getAddonByID("tbsync@jobisoft.de");
              const eas = await AddonManager.getAddonByID("eas4tbsync@jobisoft.de");

              if (!tbs || !tbs.isActive) {
                return { ok: false, error: "TbSync is not installed/enabled (expected add-on id: tbsync@jobisoft.de)." };
              }
              if (!eas || !eas.isActive) {
                return { ok: false, error: "Provider for Exchange ActiveSync is not installed/enabled (expected add-on id: eas4tbsync@jobisoft.de)." };
              }
              return { ok: true };
            }

            async function getTbSyncModule() {
              // Ensure required add-ons are installed.
              const check = await assertTbSyncInstalled();
              if (!check.ok) {
                throw new Error(check.error);
              }

              const tbsyncExtension = ExtensionParent.GlobalManager.getExtension("tbsync@jobisoft.de");
              if (!tbsyncExtension) {
                throw new Error("TbSync extension is not available via ExtensionParent.GlobalManager.");
              }

              const { TbSync: TbSyncModule } = ChromeUtils.importESModule(
                `chrome://tbsync/content/tbsync.sys.mjs?${tbsyncExtension.manifest.version}`
              );

              // TbSync may not have been loaded yet; load it if needed.
              if (!TbSyncModule.enabled) {
                const addon = await AddonManager.getAddonByID("tbsync@jobisoft.de");
                // TbSync expects its own extension object as the second argument.
                await TbSyncModule.load(addon, tbsyncExtension);
              }

              return TbSyncModule;
            }

            function toJsDate(calDateTime) {
              try {
                if (calDateTime && calDateTime.jsDate) return calDateTime.jsDate;
              } catch {}
              try {
                return cal.dtz.dateTimeToJsDate(calDateTime);
              } catch {
                return null;
              }
            }

            function clampDateTimeToRange(dt, startJs, endJs) {
              const js = toJsDate(dt);
              if (!js) return false;
              const t = js.getTime();
              return t >= startJs.getTime() && t < endJs.getTime();
            }

            async function tbsyncListEasCalendarFolders(args) {
              const { tbsyncUser, includeUnselected, types } = args || {};
              const TbSyncModule = await getTbSyncModule();
              try {
                const provider = new TbSyncModule.ProviderData("eas");
                const folderArgs = { selected: includeUnselected ? false : true };
                if (Array.isArray(types) && types.length > 0) folderArgs.type = types;
                else folderArgs.type = ["8", "13"]; // default: calendars

                const folders = provider.getFolders(folderArgs);
                return folders
                  .filter((f) => {
                    if (!tbsyncUser) return true;
                    const u = String(f.accountData.getAccountProperty("user") || "");
                    return u === String(tbsyncUser);
                  })
                  .map((f) => ({
                    accountID: String(f.accountID),
                    folderID: String(f.folderID),
                    type: String(f.getFolderProperty("type") || ""),
                    serverID: String(f.getFolderProperty("serverID") || ""),
                    foldername: String(f.getFolderProperty("foldername") || ""),
                    selected: !!f.getFolderProperty("selected"),
                    targetCalendarId: String(f.getFolderProperty("target") || ""),
                    targetName: String(f.getFolderProperty("targetName") || ""),
                  }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function tbsyncSelectEasCalendarFolder(args) {
              const { tbsyncUser, serverID, foldername } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!serverID && !foldername) return { error: "Provide serverID or foldername" };

              const TbSyncModule = await getTbSyncModule();
              try {
                const provider = new TbSyncModule.ProviderData("eas");
                const folders = provider.getFolders({ selected: false, type: ["8", "13"] });

                const candidates = folders.filter((f) => {
                  const u = String(f.accountData.getAccountProperty("user") || "");
                  if (u !== String(tbsyncUser)) return false;
                  if (serverID && String(f.getFolderProperty("serverID") || "") !== String(serverID)) return false;
                  if (foldername && String(f.getFolderProperty("foldername") || "") !== String(foldername)) return false;
                  return true;
                });

                if (candidates.length === 0) return { error: "No matching folder found" };

                const f = candidates[0];
                f.setFolderProperty("selected", true);
                f.setFolderProperty("status", "aborted");
                f.accountData.setAccountProperty("status", "notsyncronized");
                Services.obs.notifyObservers(null, "tbsync.observer.manager.updateSyncstate", f.accountID);

                let targetId = null;
                let targetName = null;
                const tbCalendar = await f.targetData.getTarget();
                targetId = String(tbCalendar.calendar.id || "");
                targetName = String(tbCalendar.calendar.name || "");

                try { tbsyncSyncAccountByUser(tbsyncUser); } catch {}

                return {
                  success: true,
                  accountID: String(f.accountID),
                  serverID: String(f.getFolderProperty("serverID") || ""),
                  foldername: String(f.getFolderProperty("foldername") || ""),
                  targetCalendarId: targetId,
                  targetName,
                  note: "Folder selected and target calendar ensured. Sync triggered once.",
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function tbsyncDeleteEasEventsByTitleRange(args) {
              // Best-effort delete by enumerating items through TbSync lightning wrapper.
              // This may be unreliable depending on provider behavior.
              const { calendarId, calendarName, start, end, titleMustContainAll, tbsyncUser } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const targetCal = resolveCalendar({ calendarId, calendarName });
              if (!targetCal) return { error: "Calendar not found" };

              const TbSyncModule = await getTbSyncModule();
              const provider = new TbSyncModule.ProviderData("eas");
              const folders = provider.getFolders({ selected: true, type: ["8", "13"] });
              const matchingFolders = folders.filter((f) => String(f.getFolderProperty("target") || "") === String(targetCal.id));

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              const startJs = toJsDate(range.start);
              const endJs = toJsDate(range.end);

              let deleted = 0;
              let matched = 0;
              let totalItems = 0;
              let errors = [];

              for (const folder of matchingFolders) {
                try {
                  const tbCalendar = await folder.targetData.getTarget();
                  const items = await tbCalendar.getAllItems();
                  if (Array.isArray(items)) totalItems += items.length;

                  if (!Array.isArray(items)) {
                    errors.push("getAllItems did not return an array");
                    continue;
                  }

                  for (const nativeItem of items) {
                    const title = nativeItem && nativeItem.title ? nativeItem.title : "";
                    if (!containsAll(title, titleMustContainAll)) continue;
                    if (!clampDateTimeToRange(nativeItem.startDate, startJs, endJs)) continue;

                    matched++;
                    const tbItem = new TbSyncModule.lightning.TbItem(tbCalendar, nativeItem);
                    await tbCalendar.deleteItem(tbItem, false);
                    deleted++;
                  }
                } catch (e) {
                  errors.push(e.toString());
                }
              }

              try {
                tbsyncSyncAccountByUser(tbsyncUser);
                tbsyncSyncAccountByUser(tbsyncUser);
              } catch {}

              return { success: true, folderCount: matchingFolders.length, totalItems, matched, deleted, errors };
            }

            async function tbsyncDeleteEasEventsBySqlChangelog(args) {
              // Legacy approach (manual changelog writes). Kept for reference, but NOT recommended.
              // It can fail when TbSync expects different parentId/itemId formatting.
              const { tbsyncUser, calendarId, calendarName, start, end, titleMustContainAll } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const targetCal = resolveCalendar({ calendarId, calendarName });
              if (!targetCal) return { error: "Calendar not found" };

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              const TbSyncModule = await getTbSyncModule();
              const { Sqlite } = ChromeUtils.importESModule("resource://gre/modules/Sqlite.sys.mjs");

              const prof = Services.dirsvc.get("ProfD", Ci.nsIFile);
              prof.append("calendar-data");
              prof.append("local.sqlite");
              const dbPath = prof.path;

              const conn = await Sqlite.openConnection({ path: dbPath });
              let ids = [];
              try {
                const params = {
                  cal_id: targetCal.id,
                  start_us: range.start.nativeTime,
                  end_us: range.end.nativeTime,
                };

                const clauses = ["cal_id = :cal_id", "event_start >= :start_us AND event_start < :end_us"]; 
                let idx = 0;
                for (const kwRaw of (titleMustContainAll || [])) {
                  const kw = String(kwRaw || "").toLowerCase().trim();
                  if (!kw) continue;
                  idx += 1;
                  const key = `kw${idx}`;
                  params[key] = `%${kw}%`;
                  clauses.push(`lower(title) LIKE :${key}`);
                }

                const sql = `SELECT id, title FROM cal_events WHERE ${clauses.join(" AND ")}`;
                const rows = await conn.execute(sql, params);
                ids = rows.map((r) => ({ id: r.getResultByName("id"), title: r.getResultByName("title") }));
              } finally {
                await conn.close();
              }

              for (const it of ids) {
                try {
                  TbSyncModule.db.addItemToChangeLog(targetCal.id, String(it.id), "deleted_by_user");
                } catch {}
              }

              try {
                tbsyncSyncAccountByUser(tbsyncUser);
              } catch {}

              return { success: true, calendarId: targetCal.id, candidateCount: ids.length, sample: ids.slice(0, 10), note: "Legacy changelog-based delete; may not be processed by TbSync." };
            }

            async function tbsyncDeleteEasEventsNativeBySql(args) {
              // Robust delete for EAS calendars:
              // SQL find -> TbSync TbCalendar.getItem -> set deleted_by_user -> TbCalendar.deleteItem -> sync
              const { tbsyncUser, calendarId, calendarName, start, end, titleMustContainAll } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const targetCal = resolveCalendar({ calendarId, calendarName });
              if (!targetCal) return { error: "Calendar not found" };

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              const TbSyncModule = await getTbSyncModule();

              const provider = new TbSyncModule.ProviderData("eas");
              const folders = provider.getFolders({ selected: true, type: ["8", "13"] });
              const matchingFolders = folders.filter((f) => String(f.getFolderProperty("target") || "") === String(targetCal.id));
              if (matchingFolders.length === 0) return { error: "No selected EAS folder mapped to this calendar" };

              const { Sqlite } = ChromeUtils.importESModule("resource://gre/modules/Sqlite.sys.mjs");
              const prof = Services.dirsvc.get("ProfD", Ci.nsIFile);
              prof.append("calendar-data");
              prof.append("local.sqlite");
              const dbPath = prof.path;

              const conn = await Sqlite.openConnection({ path: dbPath });
              let ids = [];
              try {
                const params = {
                  cal_id: targetCal.id,
                  start_us: range.start.nativeTime,
                  end_us: range.end.nativeTime,
                };
                const clauses = ["cal_id = :cal_id", "event_start >= :start_us AND event_start < :end_us"]; 
                let idx = 0;
                for (const kwRaw of (titleMustContainAll || [])) {
                  const kw = String(kwRaw || "").toLowerCase().trim();
                  if (!kw) continue;
                  idx += 1;
                  const key = `kw${idx}`;
                  params[key] = `%${kw}%`;
                  clauses.push(`lower(title) LIKE :${key}`);
                }
                const sql = `SELECT id, title FROM cal_events WHERE ${clauses.join(" AND ")}`;
                const rows = await conn.execute(sql, params);
                ids = rows.map((r) => ({ id: String(r.getResultByName("id")), title: String(r.getResultByName("title") || "") }));
              } finally {
                await conn.close();
              }

              let deleted = 0;
              let notFound = 0;
              let errors = [];

              for (const folder of matchingFolders) {
                let tbCalendar;
                try {
                  tbCalendar = await folder.targetData.getTarget();
                } catch (e) {
                  errors.push(e.toString());
                  continue;
                }

                for (const it of ids) {
                  try {
                    const tbItem = await Promise.race([
                      tbCalendar.getItem(it.id),
                      timeoutAfter(3000).then(() => null),
                    ]);
                    if (!tbItem) {
                      notFound += 1;
                      continue;
                    }
                    tbItem.changelogStatus = "deleted_by_user";
                    await tbCalendar.deleteItem(tbItem, false);
                    deleted += 1;
                  } catch (e) {
                    errors.push(e.toString());
                  }
                }
              }

              try {
                tbsyncSyncAccountByUser(tbsyncUser);
              } catch {}

              return { success: true, calendarId: targetCal.id, candidateCount: ids.length, deleted, notFound, errors: errors.slice(0, 5), sample: ids.slice(0, 10) };
            }

            async function tbsyncCreateEasEventNative(args) {
              // Create an event via TbSync lightning wrapper (ensures _by_user changelog) then sync.
              const {
                tbsyncUser,
                calendarId,
                calendarName,
                title,
                description,
                location,
                allDay,
                date,
                start,
                end,
                mcpTag,
              } = args || {};

              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!title) return { error: "Missing required field: title" };

              const targetCal = resolveCalendar({ calendarId, calendarName });
              if (!targetCal) return { error: "Calendar not found" };

              const TbSyncModule = await getTbSyncModule();
              const provider = new TbSyncModule.ProviderData("eas");
              const folders = provider.getFolders({ selected: true, type: ["8", "13"] });
              const matchingFolders = folders.filter((f) => String(f.getFolderProperty("target") || "") === String(targetCal.id));
              if (matchingFolders.length === 0) return { error: "No selected EAS folder mapped to this calendar" };

              // Build the Lightning calEvent using existing helpers.
              let created = 0;
              let errors = [];
              let itemId = null;

              for (const folder of matchingFolders) {
                let tbCalendar;
                try {
                  tbCalendar = await folder.targetData.getTarget();
                } catch (e) {
                  errors.push(e.toString());
                  continue;
                }

                try {
                  if (typeof allDay !== "boolean") throw new Error("Missing required field: allDay (boolean)");

                  const tag = mcpTag ? String(mcpTag).trim() : shortTagForEvent({ allDay, date, start, end });

                  const ev = Cc["@mozilla.org/calendar/event;1"].createInstance(Ci.calIEvent);
                  ev.title = withTagAtEnd(title, tag);
                  try { ev.setProperty("X-MCP-TAG", tag); } catch {}
                  if (description) {
                    try { ev.setProperty("DESCRIPTION", description); } catch {}
                  }
                  if (location) {
                    try { ev.setProperty("LOCATION", location); } catch {}
                  }

                  let range = null;
                  if (allDay) {
                    if (!date) throw new Error("Missing required field for all-day events: date (YYYY-MM-DD)");
                    range = makeAllDayRange(date);
                    if (!range) throw new Error(`Invalid date format (expected YYYY-MM-DD): ${date}`);
                  } else {
                    if (!start || !end) throw new Error("Missing required fields for timed events: start, end (YYYY-MM-DDTHH:MM)");
                    range = makeTimedRange(start, end);
                    if (!range) throw new Error(`Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}`);
                    if (range.error) throw new Error(range.error);
                  }

                  ev.startDate = range.start;
                  ev.endDate = range.end;

                  const tbItem = new TbSyncModule.lightning.TbItem(tbCalendar, ev);
                  tbItem.changelogStatus = "added_by_user";
                  await tbCalendar.addItem(tbItem, false);
                  itemId = String(ev.id || "") || null;
                  created += 1;
                } catch (e) {
                  errors.push(e.toString());
                }
              }

              try {
                tbsyncSyncAccountByUser(tbsyncUser);
              } catch {}

              return { success: true, calendarId: targetCal.id, created, itemId, errors: errors.slice(0, 5) };
            }

            async function tbsyncModifyEasEventsBySql(args) {
              // Modify events (move/reschedule) without delete+recreate.
              // We find candidate ids via local.sqlite, then modify via TbSync lightning wrapper (modified_by_user), then sync.
              const { tbsyncUser, calendarId, calendarName, start, end, titleMustContainAll, allDay, newDate, newStart, newEnd } = args || {};
              if (!tbsyncUser) return { error: "Missing required field: tbsyncUser" };
              if (!start || !end) return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };

              const targetCal = resolveCalendar({ calendarId, calendarName });
              if (!targetCal) return { error: "Calendar not found" };

              const range = makeTimedRange(start, end);
              if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              if (range.error) return { error: range.error };

              let newRange = null;
              if (allDay) {
                if (!newDate) return { error: "Missing required field for allDay modify: newDate (YYYY-MM-DD)" };
                newRange = makeAllDayRange(newDate);
                if (!newRange) return { error: `Invalid newDate format (expected YYYY-MM-DD): ${newDate}` };
              } else {
                if (!newStart || !newEnd) return { error: "Missing required fields for timed modify: newStart, newEnd (YYYY-MM-DDTHH:MM)" };
                newRange = makeTimedRange(newStart, newEnd);
                if (!newRange) return { error: `Invalid datetime format for newStart/newEnd: ${newStart} ${newEnd}` };
                if (newRange.error) return { error: newRange.error };
              }

              const TbSyncModule = await getTbSyncModule();
              const provider = new TbSyncModule.ProviderData("eas");
              const folders = provider.getFolders({ selected: true, type: ["8", "13"] });
              const matchingFolders = folders.filter((f) => String(f.getFolderProperty("target") || "") === String(targetCal.id));
              if (matchingFolders.length === 0) return { error: "No EAS folder mapped to this calendar" };

              const { Sqlite } = ChromeUtils.importESModule("resource://gre/modules/Sqlite.sys.mjs");
              const prof = Services.dirsvc.get("ProfD", Ci.nsIFile);
              prof.append("calendar-data");
              prof.append("local.sqlite");
              const dbPath = prof.path;

              const conn = await Sqlite.openConnection({ path: dbPath });
              let ids = [];
              try {
                const params = {
                  cal_id: targetCal.id,
                  start_us: range.start.nativeTime,
                  end_us: range.end.nativeTime,
                };

                const clauses = ["cal_id = :cal_id", "event_start >= :start_us AND event_start < :end_us"]; 
                let idx = 0;
                for (const kwRaw of (titleMustContainAll || [])) {
                  const kw = String(kwRaw || "").toLowerCase().trim();
                  if (!kw) continue;
                  idx += 1;
                  const key = `kw${idx}`;
                  params[key] = `%${kw}%`;
                  clauses.push(`lower(title) LIKE :${key}`);
                }

                const sql = `SELECT id, title FROM cal_events WHERE ${clauses.join(" AND ")}`;
                const rows = await conn.execute(sql, params);
                ids = rows.map((r) => ({ id: r.getResultByName("id"), title: r.getResultByName("title") }));
              } finally {
                await conn.close();
              }

              let modified = 0;
              let errors = [];

              for (const folder of matchingFolders) {
                let tbCalendar;
                try {
                  tbCalendar = await folder.targetData.getTarget();
                } catch (e) {
                  errors.push(e.toString());
                  continue;
                }

                for (const it of ids) {
                  try {
                    const tbOld = await tbCalendar.getItem(String(it.id));
                    if (!tbOld) continue;
                    const newNative = tbOld._item.clone();
                    newNative.startDate = newRange.start;
                    newNative.endDate = newRange.end;
                    const tbNew = new TbSyncModule.lightning.TbItem(tbCalendar, newNative);
                    await tbCalendar.modifyItem(tbNew, tbOld, false);
                    modified += 1;
                  } catch (e) {
                    errors.push(e.toString());
                  }
                }
              }

              // Trigger sync (fire-and-forget)
              try {
                tbsyncSyncAccountByUser(tbsyncUser);
                tbsyncSyncAccountByUser(tbsyncUser);
              } catch {}

              return { success: true, candidateCount: ids.length, modified, errors: errors.slice(0, 5) };
            }

            async function tbsyncListAccounts() {
              const TbSyncModule = await getTbSyncModule();
              const accounts = TbSyncModule.db.getAccounts();
              return accounts.IDs.map((id) => {
                const data = accounts.data[id];
                return {
                  accountID: id,
                  accountname: data.accountname || null,
                  provider: data.provider || null,
                  user: data.user || null,
                  host: data.host || null,
                  lastsynctime: data.lastsynctime || null,
                };
              });
            }

            async function tbsyncSyncAccount(accountID) {
              const TbSyncModule = await getTbSyncModule();

              // Sync is async in effect; TbSync.core.syncAccount kicks off network ops.
              // We return immediately after triggering.
              TbSyncModule.core.syncAccount(String(accountID));
              return { success: true, accountID: String(accountID), triggered: true };
            }

            async function tbsyncSyncAccountByUser(userEmail) {
              const TbSyncModule = await getTbSyncModule();
              const accounts = TbSyncModule.db.getAccounts();
              const needle = (userEmail || "").toLowerCase().trim();
              const matches = accounts.IDs.filter((id) => {
                const data = accounts.data[id];
                return (data.user || "").toLowerCase().trim() === needle;
              });

              if (matches.length === 0) {
                return { error: `No TbSync account found with user=${needle}` };
              }
              if (matches.length > 1) {
                return { error: `Multiple TbSync accounts found with user=${needle}: ${matches.join(",")}` };
              }

              TbSyncModule.core.syncAccount(matches[0]);
              return { success: true, accountID: matches[0], user: needle, triggered: true };
            }

            async function tbsyncResetEasFolderSync(userEmail) {
              // Force EAS FolderSync to re-run from scratch by clearing the foldersynckey.
              // This is needed if TbSync's folder list got corrupted/missing the primary Calendar.
              const TbSyncModule = await getTbSyncModule();
              const accounts = TbSyncModule.db.getAccounts();
              const needle = (userEmail || "").toLowerCase().trim();
              const matches = accounts.IDs.filter((id) => {
                const data = accounts.data[id];
                return (data.user || "").toLowerCase().trim() === needle && (data.provider || "") === "eas";
              });
              if (matches.length === 0) return { error: `No EAS TbSync account found with user=${needle}` };
              if (matches.length > 1) return { error: `Multiple EAS TbSync accounts found with user=${needle}: ${matches.join(",")}` };

              const accountID = matches[0];
              try {
                TbSyncModule.db.setAccountProperty(accountID, "foldersynckey", "");
                TbSyncModule.db.setAccountProperty(accountID, "status", "notsyncronized");
              } catch (e) {
                return { error: `Failed to reset foldersynckey: ${e}` };
              }

              TbSyncModule.core.syncAccount(accountID);
              return { success: true, accountID, user: needle, reset: true, triggered: true };
            }

            function listCalendars() {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                return cal.manager.getCalendars().map((c) => ({
                  id: c.id,
                  name: c.name,
                  type: c.type,
                  readOnly: c.readOnly,
                }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function listEventsOnDate(args) {
              if (!cal) {
                return { error: "Calendar not available" };
              }

              const { calendarName, calendarId, date } = args || {};
              if (!date) return { error: "Missing required field: date (YYYY-MM-DD)" };

              const calendars = cal.manager.getCalendars();
              let target = null;
              if (calendarId) {
                target = calendars.find((c) => c.id === calendarId) || null;
              } else if (calendarName) {
                target = calendars.find((c) => c.name === calendarName) || null;
              }
              if (!target) {
                return { error: `Calendar not found (name=${calendarName || ""}, id=${calendarId || ""})` };
              }

              const [y, m, d] = String(date).split("-").map((x) => parseInt(x, 10));
              if (!y || !m || !d) {
                return { error: `Invalid date format (expected YYYY-MM-DD): ${date}` };
              }

              const start = cal.createDateTime();
              start.year = y;
              start.month = m - 1;
              start.day = d;
              start.isDate = true;
              start.timezone = cal.dtz.floating;

              const end = start.clone();
              end.day = end.day + 1;
              end.isDate = true;
              end.timezone = cal.dtz.floating;

              const FILTER = Ci.calICalendar;
              const filter =
                FILTER.ITEM_FILTER_TYPE_EVENT |
                FILTER.ITEM_FILTER_CLASS_OCCURRENCES |
                FILTER.ITEM_FILTER_COMPLETED_YES |
                FILTER.ITEM_FILTER_COMPLETED_NO;

              const job = startGetItemsJob(target, filter, start, end, {
                kind: "date",
                date,
                calendar: { id: target.id, name: target.name },
              });

              return {
                success: true,
                calendar: { id: target.id, name: target.name },
                date,
                pending: job.status === "pending",
                jobId: job.jobId,
                count: job.events.length,
                events: job.events,
                note: "Non-blocking query started. Poll with getCalendarJob(jobId) to retrieve results (and partial results while pending).",
              };
            }

            async function listCalendarEvents(args) {
              if (!cal) {
                return { error: "Calendar not available" };
              }

              const { calendarName, calendarId, start, end } = args || {};
              if (!start || !end) {
                return { error: "Missing required fields: start, end (YYYY-MM-DDTHH:MM)" };
              }

              const calendars = cal.manager.getCalendars();
              let target = null;
              if (calendarId) {
                target = calendars.find((c) => c.id === calendarId) || null;
              } else if (calendarName) {
                target = calendars.find((c) => c.name === calendarName) || null;
              }
              if (!target) {
                return { error: `Calendar not found (name=${calendarName || ""}, id=${calendarId || ""})` };
              }

              const range = makeTimedRange(start, end);
              if (!range) {
                return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
              }
              if (range.error) return { error: range.error };

              const FILTER = Ci.calICalendar;
              const filter =
                FILTER.ITEM_FILTER_TYPE_EVENT |
                FILTER.ITEM_FILTER_CLASS_OCCURRENCES |
                FILTER.ITEM_FILTER_COMPLETED_YES |
                FILTER.ITEM_FILTER_COMPLETED_NO;

              const job = startGetItemsJob(target, filter, range.start, range.end, {
                kind: "range",
                start,
                end,
                calendar: { id: target.id, name: target.name },
              });

              return {
                success: true,
                calendar: { id: target.id, name: target.name },
                start,
                end,
                pending: job.status === "pending",
                jobId: job.jobId,
                count: job.events.length,
                events: job.events,
                note: "Non-blocking query started. Poll with getCalendarJob(jobId) to retrieve results (and partial results while pending).",
              };
            }

            function parseDateYMD(dateStr) {
              const [y, m, d] = String(dateStr).split("-").map((x) => parseInt(x, 10));
              if (!y || !m || !d) return null;
              return { y, m, d };
            }

            function parseDateTimeLocal(dtStr) {
              // Accept "YYYY-MM-DDTHH:MM" (preferred) or "YYYY-MM-DD HH:MM".
              const s = String(dtStr || "").trim();
              const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2})$/);
              if (!m) return null;
              return {
                y: parseInt(m[1], 10),
                mo: parseInt(m[2], 10),
                d: parseInt(m[3], 10),
                hh: parseInt(m[4], 10),
                mm: parseInt(m[5], 10),
              };
            }

            function makeAllDayRange(dateStr) {
              const ymd = parseDateYMD(dateStr);
              if (!ymd) return null;

              const start = cal.createDateTime();
              start.year = ymd.y;
              start.month = ymd.m - 1;
              start.day = ymd.d;
              start.isDate = true;
              start.timezone = cal.dtz.floating;

              const end = start.clone();
              end.day = end.day + 1;
              end.isDate = true;
              end.timezone = cal.dtz.floating;

              return { start, end };
            }

            function makeTimedRange(startStr, endStr) {
              const s = parseDateTimeLocal(startStr);
              const e = parseDateTimeLocal(endStr);
              if (!s || !e) return null;

              const start = cal.createDateTime();
              start.year = s.y;
              start.month = s.mo - 1;
              start.day = s.d;
              start.hour = s.hh;
              start.minute = s.mm;
              start.second = 0;
              start.isDate = false;
              // Use default timezone so events display correctly across clients.
              start.timezone = cal.dtz.defaultTimezone;

              const end = cal.createDateTime();
              end.year = e.y;
              end.month = e.mo - 1;
              end.day = e.d;
              end.hour = e.hh;
              end.minute = e.mm;
              end.second = 0;
              end.isDate = false;
              end.timezone = cal.dtz.defaultTimezone;

              if (end.compare(start) <= 0) return { error: "end must be after start" };

              return { start, end };
            }

            function resolveCalendar(args) {
              const { calendarName, calendarId } = args || {};
              const calendars = cal.manager.getCalendars();
              let target = null;
              if (calendarId) {
                target = calendars.find((c) => c.id === calendarId) || null;
              } else if (calendarName) {
                target = calendars.find((c) => c.name === calendarName) || null;
              }
              return target;
            }

            function maybeRefreshCalendar(target) {
              try {
                if (target && typeof target.refresh === "function") {
                  target.refresh();
                  return { refreshed: true };
                }
              } catch (e) {
                return { refreshed: false, error: e.toString() };
              }
              return { refreshed: false };
            }

            function shortTagForEvent(args) {
              const { allDay, date, start, end } = args || {};
              if (allDay && date) return `mcp:${String(date).replace(/-/g, "")}`; // mcp:20260204
              if (!allDay && start && end) {
                const s = String(start).replace(/[-:T\s]/g, ""); // 20260204HHMM
                const e = String(end).replace(/[-:T\s]/g, "");
                // keep short: date+start-end times
                return `mcp:${s.slice(0, 8)}-${s.slice(8, 12)}-${e.slice(8, 12)}`; // mcp:20260204-0600-0700
              }
              return "mcp";
            }

            function withTagAtEnd(title, tag) {
              const base = String(title || "").trim();
              const t = String(tag || "").trim();
              if (!t) return base;
              // Put tag at the very end, short.
              if (base.endsWith(`(${t})`)) return base;
              return `${base} (${t})`;
            }

            async function createCalendarEvent(args) {
              if (!cal) {
                return { error: "Calendar not available" };
              }

              const {
                title,
                description,
                allDay,
                // All-day:
                date,
                // Timed:
                start,
                end,
                // Optional short tag placed at end of title.
                mcpTag,
              } = args || {};

              if (!title) return { error: "Missing required field: title" };
              if (typeof allDay !== "boolean") return { error: "Missing required field: allDay (boolean)" };

              const target = resolveCalendar(args);
              if (!target) {
                return { error: `Calendar not found (name=${(args || {}).calendarName || ""}, id=${(args || {}).calendarId || ""})` };
              }
              if (target.readOnly) {
                return { error: `Calendar is read-only: ${target.name}` };
              }

              // Some providers won't return results or complete operations until the calendar is refreshed.
              maybeRefreshCalendar(target);

              const ev = Cc["@mozilla.org/calendar/event;1"].createInstance(Ci.calIEvent);
              const tag = mcpTag ? String(mcpTag).trim() : shortTagForEvent({ allDay, date, start, end });
              ev.title = withTagAtEnd(title, tag);
              // Also store tag in a property for debugging/search in raw iCal.
              try {
                ev.setProperty("X-MCP-TAG", tag);
              } catch {}
              if (description) {
                ev.setProperty("DESCRIPTION", description);
              }

              let range = null;
              if (allDay) {
                if (!date) return { error: "Missing required field for all-day events: date (YYYY-MM-DD)" };
                range = makeAllDayRange(date);
                if (!range) return { error: `Invalid date format (expected YYYY-MM-DD): ${date}` };
              } else {
                if (!start || !end) return { error: "Missing required fields for timed events: start, end (YYYY-MM-DDTHH:MM)" };
                range = makeTimedRange(start, end);
                if (!range) return { error: `Invalid datetime format (expected YYYY-MM-DDTHH:MM): start=${start} end=${end}` };
                if (range.error) return { error: range.error };
              }

              ev.startDate = range.start;
              ev.endDate = range.end;

              const job = startAddItemJob(target, ev, {
                kind: "addItem",
                calendar: { id: target.id, name: target.name },
                title,
                allDay,
                date: allDay ? date : undefined,
                start: !allDay ? range.start.icalString : undefined,
                end: !allDay ? range.end.icalString : undefined,
              });

              return {
                success: true,
                calendar: { id: target.id, name: target.name },
                title,
                allDay,
                date: allDay ? date : undefined,
                start: !allDay ? range.start.icalString : undefined,
                end: !allDay ? range.end.icalString : undefined,
                pending: job.status === "pending",
                jobId: job.jobId,
                itemId: job.itemId || null,
                note: "Non-blocking create started. Poll with getCalendarJob(jobId) to confirm completion and retrieve itemId.",
              };
            }

            async function sleepMs(ms) {
              await timeoutAfter(ms);
            }

            function cleanupIdempotency() {
              const now = Date.now();
              for (const [k, v] of _idempotency.entries()) {
                if (!v || !v.ts || now - v.ts > IDEMPOTENCY_TTL_MS) {
                  _idempotency.delete(k);
                }
              }
            }

            function idempotencyKey(args) {
              // calendar + time/date + tag + (for all-day) title to avoid suppressing distinct all-day events on same date.
              const calKey = (args && (args.calendarId || args.calendarName)) || "";
              const tag = args && args.mcpTag ? String(args.mcpTag).trim() : shortTagForEvent(args);
              if (args && args.allDay) {
                const dateKey = String(args.date || "");
                const titleKey = String(args.title || "");
                return calKey + '|' + dateKey + '|' + tag + '|' + titleKey;
              }
              const timeKey = String(args.start || "") + '-' + String(args.end || "");
              return calKey + '|' + timeKey + '|' + tag;
            }

            async function syncAndCreateCalendarEvent(args) {
              // Non-blocking robust workflow.
              // Returns immediately with a workflowJobId.
              const { tbsyncUser } = args || {};
              if (!tbsyncUser) {
                return { error: "Missing required field: tbsyncUser (email)" };
              }

              cleanupIdempotency();

              const tag = (args && args.mcpTag) ? String(args.mcpTag).trim() : shortTagForEvent(args);
              const argsWithTag = Object.assign({}, args, { mcpTag: tag, title: withTagAtEnd(args.title, tag) });

              const key = idempotencyKey(argsWithTag);
              const seen = _idempotency.get(key);
              if (seen) {
                return {
                  success: true,
                  alreadyCreated: true,
                  idempotencyKey: key,
                  created: seen.created || null,
                  note: "Duplicate request suppressed (idempotency cache).",
                };
              }

              const wfJobId = newWorkflowJobId();
              const job = {
                jobId: wfJobId,
                status: "running",
                createdAtMs: Date.now(),
                updatedAtMs: Date.now(),
                step: "start",
                meta: { idempotencyKey: key, tbsyncUser, calendar: argsWithTag.calendarId || argsWithTag.calendarName || null },
                result: null,
                error: null,
              };
              _workflowJobs.set(wfJobId, job);

              // Mark idempotency immediately to avoid duplicates while the workflow runs.
              _idempotency.set(key, { ts: Date.now(), created: { pending: true, workflowJobId: wfJobId } });

              (async () => {
                try {
                  job.step = "preSync";
                  job.updatedAtMs = Date.now();
                  const pre = await tbsyncSyncAccountByUser(tbsyncUser);

                  job.step = "create";
                  job.updatedAtMs = Date.now();
                  await sleepMs(1500);
                  const created = await createCalendarEvent(argsWithTag);

                  job.step = "postSync";
                  job.updatedAtMs = Date.now();
                  await sleepMs(1500);
                  const post = await tbsyncSyncAccountByUser(tbsyncUser);
                  await sleepMs(1500);
                  const post2 = await tbsyncSyncAccountByUser(tbsyncUser);

                  job.status = "done";
                  job.step = "done";
                  job.updatedAtMs = Date.now();
                  job.result = { preSync: pre, created, postSync: post, postSync2: post2 };

                  // Update idempotency cache with actual created payload.
                  _idempotency.set(key, { ts: Date.now(), created: { preSync: pre, created, postSync: post, postSync2: post2 } });
                } catch (e) {
                  job.status = "error";
                  job.step = "error";
                  job.updatedAtMs = Date.now();
                  job.error = e.toString();
                }
              })();

              return {
                success: true,
                pending: true,
                workflowJobId: wfJobId,
                idempotencyKey: key,
                note: "Workflow started (sync→create→sync). Poll with getWorkflowJob(workflowJobId).",
              };
            }

            const tools = [
              {
                name: "tbsyncListAccounts",
                title: "TbSync List Accounts",
                description: "List TbSync accounts (requires TbSync + EAS provider installed)",
                inputSchema: { type: "object", properties: {}, required: [] },
              },
              {
                name: "tbsyncSyncAccount",
                title: "TbSync Sync Account",
                description: "Trigger a TbSync sync for a specific TbSync account ID",
                inputSchema: {
                  type: "object",
                  properties: { accountID: { type: "string" } },
                  required: ["accountID"],
                },
              },
              {
                name: "tbsyncSyncAccountByUser",
                title: "TbSync Sync Account (by user email)",
                description: "Trigger a TbSync sync by matching TbSync account user (email)",
                inputSchema: {
                  type: "object",
                  properties: { userEmail: { type: "string" } },
                  required: ["userEmail"],
                },
              },
              {
                name: "tbsyncResetEasFolderSync",
                title: "TbSync Reset EAS FolderSync",
                description: "Clear EAS foldersynckey to force refetch of folder hierarchy (repairs missing primary Calendar folder)",
                inputSchema: {
                  type: "object",
                  properties: { userEmail: { type: "string" } },
                  required: ["userEmail"],
                },
              },
              {
                name: "listCalendars",
                title: "List Calendars",
                description: "List Thunderbird calendars",
                inputSchema: { type: "object", properties: {}, required: [] },
              },
              {
                name: "tbsyncListEasCalendarFolders",
                title: "TbSync List EAS Calendar Folders",
                description: "List EAS calendar folders (selected by default; set includeUnselected=true to list all)",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    includeUnselected: { type: "boolean" },
                    types: { type: "array", items: { type: "string" } }
                  },
                  required: [],
                },
              },
              {
                name: "tbsyncSelectEasCalendarFolder",
                title: "TbSync Select EAS Calendar Folder",
                description: "Select an EAS calendar folder (subscribe) and ensure its Thunderbird calendar target exists, then sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    serverID: { type: "string" },
                    foldername: { type: "string" }
                  },
                  required: ["tbsyncUser"],
                },
              },
              {
                name: "tbsyncDeleteEasEventsByTitleRange",
                title: "TbSync Delete EAS Events By Title/Range",
                description: "(Best-effort) Enumerate EAS calendar items and delete matches as user deletes, then TbSync sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string" },
                    end: { type: "string" },
                    titleMustContainAll: { type: "array", items: { type: "string" } }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll"],
                },
              },
              {
                name: "tbsyncDeleteEasEventsBySqlChangelog",
                title: "TbSync Delete EAS Events (SQL→Changelog)",
                description: "Legacy: find matches via local.sqlite, then write deleted_by_user directly to TbSync changelog (not recommended)",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string" },
                    end: { type: "string" },
                    titleMustContainAll: { type: "array", items: { type: "string" } }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll"],
                },
              },
              {
                name: "tbsyncDeleteEasEventsNativeBySql",
                title: "TbSync Delete EAS Events (SQL→TbSync deleteItem)",
                description: "Robust: find matches via local.sqlite, then delete via TbSync lightning wrapper as user deletes and sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string" },
                    end: { type: "string" },
                    titleMustContainAll: { type: "array", items: { type: "string" } }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll"],
                },
              },
              {
                name: "tbsyncCreateEasEventNative",
                title: "TbSync Create EAS Event (native addItem)",
                description: "Robust: create event through TbSync lightning wrapper as user adds and sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    title: { type: "string" },
                    description: { type: "string" },
                    location: { type: "string" },
                    allDay: { type: "boolean" },
                    date: { type: "string" },
                    start: { type: "string" },
                    end: { type: "string" },
                    mcpTag: { type: "string" }
                  },
                  required: ["tbsyncUser", "title", "allDay"],
                },
              },
              {
                name: "tbsyncModifyEasEventsBySql",
                title: "TbSync Modify EAS Events (SQL→modifyItem)",
                description: "Modify/move events without delete+recreate: find candidates via local.sqlite, then modifyItem as user and sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string" },
                    end: { type: "string" },
                    titleMustContainAll: { type: "array", items: { type: "string" } },
                    allDay: { type: "boolean" },
                    newDate: { type: "string" },
                    newStart: { type: "string" },
                    newEnd: { type: "string" }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll", "allDay"],
                },
              },
              {
                name: "createCalendarEvent",
                title: "Create Calendar Event",
                description: "Create an event in a calendar (all-day or timed)",
                inputSchema: {
                  type: "object",
                  properties: {
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    title: { type: "string" },
                    description: { type: "string" },
                    allDay: { type: "boolean" },
                    date: { type: "string", description: "YYYY-MM-DD (required when allDay=true)" },
                    start: { type: "string", description: "YYYY-MM-DDTHH:MM (required when allDay=false)" },
                    end: { type: "string", description: "YYYY-MM-DDTHH:MM (required when allDay=false)" },
                    mcpTag: { type: "string", description: "Optional short tag appended to title (placed at end)." }
                  },
                  required: ["title", "allDay"],
                },
              },
              {
                name: "listEventsOnDate",
                title: "List Events on Date",
                description: "List events on a specific date in a calendar (for verification/debug)",
                inputSchema: {
                  type: "object",
                  properties: {
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    date: { type: "string", description: "YYYY-MM-DD" }
                  },
                  required: ["date"],
                },
              },
              {
                name: "listCalendarEvents",
                title: "List Calendar Events",
                description: "Start a non-blocking query to list events in a calendar over a date/time range (poll via getCalendarJob)",
                inputSchema: {
                  type: "object",
                  properties: {
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string", description: "YYYY-MM-DDTHH:MM" },
                    end: { type: "string", description: "YYYY-MM-DDTHH:MM" }
                  },
                  required: ["start", "end"],
                },
              },
              {
                name: "getCalendarJob",
                title: "Get Calendar Job",
                description: "Poll a non-blocking calendar query job by jobId",
                inputSchema: {
                  type: "object",
                  properties: {
                    jobId: { type: "string" }
                  },
                  required: ["jobId"],
                },
              },
              {
                name: "getWorkflowJob",
                title: "Get Workflow Job",
                description: "Poll a non-blocking sync→create→sync workflow job",
                inputSchema: {
                  type: "object",
                  properties: {
                    workflowJobId: { type: "string" }
                  },
                  required: ["workflowJobId"],
                },
              },
              {
                name: "syncAndCreateCalendarEvent",
                title: "Sync + Create Calendar Event",
                description: "Start a non-blocking workflow: pre-sync, create event locally, post-sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string", description: "TbSync account user email (e.g., jl4624@cornell.edu)" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    title: { type: "string" },
                    description: { type: "string" },
                    allDay: { type: "boolean" },
                    date: { type: "string", description: "YYYY-MM-DD (required when allDay=true)" },
                    start: { type: "string", description: "YYYY-MM-DDTHH:MM (required when allDay=false)" },
                    end: { type: "string", description: "YYYY-MM-DDTHH:MM (required when allDay=false)" },
                    mcpTag: { type: "string", description: "Optional short tag appended to title (placed at end)." }
                  },
                  required: ["tbsyncUser", "title", "allDay"],
                },
              },
              {
                name: "syncAndDeleteEventsByTitle",
                title: "Sync + Delete Events By Title",
                description: "Legacy provider-based scan+delete (may not work for TbSync/EAS calendars)",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string", description: "YYYY-MM-DDTHH:MM" },
                    end: { type: "string", description: "YYYY-MM-DDTHH:MM" },
                    titleMustContainAll: { type: "array", items: { type: "string" }, description: "All substrings that must appear in title (case-insensitive)" }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll"],
                },
              },
              {
                name: "syncAndDeleteEventsBySql",
                title: "Sync + Delete Events (SQL match)",
                description: "Robust bulk delete: find candidates via local.sqlite, delete via provider, then TbSync sync",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    start: { type: "string", description: "YYYY-MM-DDTHH:MM" },
                    end: { type: "string", description: "YYYY-MM-DDTHH:MM" },
                    titleMustContainAll: { type: "array", items: { type: "string" }, description: "All substrings that must appear in title (case-insensitive)" }
                  },
                  required: ["tbsyncUser", "start", "end", "titleMustContainAll"],
                },
              },
            ];

            async function callTool(name, args) {
              switch (name) {
                case "tbsyncListAccounts":
                  return await tbsyncListAccounts();
                case "tbsyncSyncAccount":
                  return await tbsyncSyncAccount(args.accountID);
                case "tbsyncSyncAccountByUser":
                  return await tbsyncSyncAccountByUser(args.userEmail);
                case "tbsyncResetEasFolderSync":
                  return await tbsyncResetEasFolderSync(args.userEmail);
                case "listCalendars":
                  return listCalendars();
                case "tbsyncListEasCalendarFolders":
                  return await tbsyncListEasCalendarFolders(args);
                case "tbsyncSelectEasCalendarFolder":
                  return await tbsyncSelectEasCalendarFolder(args);
                case "tbsyncDeleteEasEventsByTitleRange":
                  return await tbsyncDeleteEasEventsByTitleRange(args);
                case "tbsyncDeleteEasEventsBySqlChangelog":
                  return await tbsyncDeleteEasEventsBySqlChangelog(args);
                case "tbsyncDeleteEasEventsNativeBySql":
                  return await tbsyncDeleteEasEventsNativeBySql(args);
                case "tbsyncCreateEasEventNative":
                  return await tbsyncCreateEasEventNative(args);
                case "tbsyncModifyEasEventsBySql":
                  return await tbsyncModifyEasEventsBySql(args);
                case "createCalendarEvent":
                  return await createCalendarEvent(args);
                case "listEventsOnDate":
                  return await listEventsOnDate(args);
                case "listCalendarEvents":
                  return await listCalendarEvents(args);
                case "getCalendarJob":
                  return getCalendarJob(args.jobId);
                case "syncAndCreateCalendarEvent":
                  return await syncAndCreateCalendarEvent(args);
                case "getWorkflowJob":
                  return getWorkflowJob(args.workflowJobId);
                case "syncAndDeleteEventsByTitle":
                  return await syncAndDeleteEventsByTitle(args);
                case "syncAndDeleteEventsBySql":
                  return await syncAndDeleteEventsBySql(args);
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.write("POST only");
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 400, "Bad Request");
                res.write("Invalid JSON");
                res.finish();
                return;
              }

              const id = message.id;
              const method = message.method;
              const params = message.params || {};

              async function reply(resultObj) {
                const payload = sanitizeForJson(JSON.stringify({ jsonrpc: "2.0", id, result: resultObj }));
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json", false);
                res.write(payload);
                res.finish();
              }

              async function replyError(errMsg) {
                const payload = sanitizeForJson(
                  JSON.stringify({
                    jsonrpc: "2.0",
                    id,
                    error: { code: -32000, message: errMsg || "Unknown error" },
                  })
                );
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json", false);
                res.write(payload);
                res.finish();
              }

              (async () => {
                try {
                  if (method === "tools/list") {
                    await reply({ tools });
                    return;
                  }

                  if (method === "tools/call") {
                    const toolName = params.name;
                    const args = params.arguments || {};
                    const out = await callTool(toolName, args);
                    await reply({ content: [{ type: "text", text: sanitizeForJson(JSON.stringify(out)) }] });
                    return;
                  }

                  await replyError(`Unknown method: ${method}`);
                } catch (e) {
                  await replyError(e.toString());
                }
              })();
            });

            server.start(MCP_PORT);
            _serverStarted = true;
            _serverStarting = false;
            _serverPort = MCP_PORT;

            return { success: true, port: MCP_PORT };
          } catch (e) {
            _serverStarting = false;
            return { success: false, error: e.toString() };
          }
        },
      },
    };
  }
};
