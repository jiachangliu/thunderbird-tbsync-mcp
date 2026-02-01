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

const MCP_PORT = 8766;

var tbsyncMcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return {
      tbsyncMcpServer: {
        start: async function () {
          try {
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

            async function createCalendarEvent(args) {
              if (!cal) {
                return { error: "Calendar not available" };
              }

              const {
                calendarName,
                calendarId,
                title,
                description,
                allDay,
                // For all-day events: use YYYY-MM-DD in local timezone (floating).
                date,
              } = args || {};

              if (!title) return { error: "Missing required field: title" };

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
              if (target.readOnly) {
                return { error: `Calendar is read-only: ${target.name}` };
              }

              if (!allDay) {
                return { error: "Only allDay=true is implemented in v0.1.0" };
              }
              if (!date) {
                return { error: "Missing required field for all-day events: date (YYYY-MM-DD)" };
              }

              // Create an all-day event.
              const ev = Cc["@mozilla.org/calendar/event;1"].createInstance(Ci.calIEvent);
              ev.title = title;
              if (description) {
                ev.setProperty("DESCRIPTION", description);
              }

              const [y, m, d] = String(date).split("-").map((x) => parseInt(x, 10));
              if (!y || !m || !d) {
                return { error: `Invalid date format (expected YYYY-MM-DD): ${date}` };
              }

              const start = cal.createDateTime();
              start.year = y;
              start.month = m - 1; // CalDateTime uses 0-based months
              start.day = d;
              start.isDate = true;
              start.timezone = cal.dtz.floating;

              const end = start.clone();
              // All-day events are typically [start, next day)
              end.day = end.day + 1;
              end.isDate = true;
              end.timezone = cal.dtz.floating;

              ev.startDate = start;
              ev.endDate = end;

              const result = await new Promise((resolve) => {
                try {
                  target.addItem(ev, {
                    onOperationComplete: function (_cal, status, opType, id, detail) {
                      resolve({ status, opType, id, detail });
                    },
                    onGetResult: function () {},
                  });
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });

              if (result.error) return result;

              return {
                success: true,
                calendar: { id: target.id, name: target.name },
                itemId: result.id || null,
                title,
                allDay: true,
                date,
              };
            }

            async function syncAndCreateCalendarEvent(args) {
              // Workflow: sync (by TbSync user email) -> create -> sync
              const { tbsyncUser } = args || {};
              if (!tbsyncUser) {
                return { error: "Missing required field: tbsyncUser (email)" };
              }

              const pre = await tbsyncSyncAccountByUser(tbsyncUser);
              if (pre.error) return pre;

              const created = await createCalendarEvent(args);
              if (created.error) return created;

              const post = await tbsyncSyncAccountByUser(tbsyncUser);
              if (post.error) {
                // Event already created locally; report post-sync error explicitly.
                return { error: post.error, created };
              }

              return { success: true, preSync: pre, created, postSync: post };
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
                name: "listCalendars",
                title: "List Calendars",
                description: "List Thunderbird calendars",
                inputSchema: { type: "object", properties: {}, required: [] },
              },
              {
                name: "createCalendarEvent",
                title: "Create Calendar Event",
                description: "Create an event in a calendar (v0.1.0: all-day only)",
                inputSchema: {
                  type: "object",
                  properties: {
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    title: { type: "string" },
                    description: { type: "string" },
                    allDay: { type: "boolean" },
                    date: { type: "string", description: "YYYY-MM-DD (for all-day events)" },
                  },
                  required: ["title", "allDay", "date"],
                },
              },
              {
                name: "syncAndCreateCalendarEvent",
                title: "Sync + Create Calendar Event",
                description: "Pre-sync, create event locally, then post-sync (safest workflow)",
                inputSchema: {
                  type: "object",
                  properties: {
                    tbsyncUser: { type: "string", description: "TbSync account user email (e.g., jl4624@cornell.edu)" },
                    calendarName: { type: "string" },
                    calendarId: { type: "string" },
                    title: { type: "string" },
                    description: { type: "string" },
                    allDay: { type: "boolean" },
                    date: { type: "string" },
                  },
                  required: ["tbsyncUser", "title", "allDay", "date"],
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
                case "listCalendars":
                  return listCalendars();
                case "createCalendarEvent":
                  return await createCalendarEvent(args);
                case "syncAndCreateCalendarEvent":
                  return await syncAndCreateCalendarEvent(args);
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

            return { success: true, port: MCP_PORT };
          } catch (e) {
            return { success: false, error: e.toString() };
          }
        },
      },
    };
  }
};
