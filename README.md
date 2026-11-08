# Thunderbird TbSync MCP

MCP-compatible local server for **calendar synchronization and calendar edits** in Thunderbird **via TbSync + EAS-4-TbSync**.

This project is intentionally **separate** from `thunderbird-mcp`:
- `thunderbird-mcp` targets **mail + lightweight calendar listing** with no dependency on TbSync.
- `thunderbird-tbsync-mcp` targets **safe calendar write workflows** that require **explicit pre-sync and post-sync** using TbSync.

## Why this exists

If you edit calendars locally while the cloud calendar has newer changes, later sync/merge can get messy.

This MCP server enforces the workflow:

1. **Sync down** from the cloud (TbSync) for the account/calendar you will modify
2. **Apply the local change** (create/edit event)
3. **Sync up** to the cloud immediately (TbSync)

The goal is to avoid divergence and reduce merge conflicts.

## Architecture

```
MCP Client <--stdio--> mcp-bridge.cjs <--HTTP--> Thunderbird Extension (localhost:8766)
```

- The Thunderbird extension runs a small local HTTP JSON-RPC server (port **8766** by default).
- The Node.js bridge provides MCP stdio protocol and forwards requests to the HTTP server.

## Requirements

- Thunderbird
- TbSync installed (`tbsync@jobisoft.de`)
- EAS-4-TbSync installed (Exchange ActiveSync provider)

If TbSync (or the EAS provider) is not installed, tools will return a clear error.

## TbSync 5.x read-only backend

The stdio bridge supports an opt-in backend for Thunderbird 153 ESR with
TbSync 5.3.x **beta** and EAS-4-TbSync 5.3.x. This backend uses TbSync's native
helper protocol 2. It does not use this project's legacy XPI or HTTP server.

In a test profile, use the TbSync Manager's Bridge tab with the matching
native helper, keep unrestricted mode off, and grant one EAS calendar.
Set `TBSYNC_MCP_BACKEND=v5` in your MCP client's environment and point the
client to `mcp-bridge.cjs` in this checkout. The helper listens on
`127.0.0.1:47654` by default. Its static bearer header protects against
simple cross-origin browser requests. Local processes can still reach
the helper while it is enabled.

- `tbsyncGetSyncStatus` reports the binding and last recorded sync, with counts
  for unpushed edits and queued calendar messages. A `null` count means the
  provider queue was unavailable. This call never requests a sync.
- `listCalendars` returns the one granted EAS calendar.
- `listCalendarEvents` takes `start` and `end` with an explicit timezone offset
  or `Z`. It returns iCalendar records and expanded recurrence instances for
  up to 31 days. The optional `calendarId` must match the grant. Results are
  capped at 50 by default and at 100 when requested, with a `truncated` flag.

V5 MCP writes and manual sync remain gated on live Outlook validation.
Use a disposable Thunderbird profile and a dedicated test calendar for
that validation. For one tagged event with no attendees, verify create,
edit, and delete in the provider queue and independently in Outlook.
Check that a repeated request creates no duplicates and that no calendar
messages were queued. The repository's normal build still packages the
legacy extension. Run `node --test` to check the v5 stdio backend offline.

## Legacy tools

Legacy mode remains the default. Its tools use JSON-RPC on
`http://localhost:8766`.

- `tbsyncListAccounts`
- `tbsyncSyncAccount` (by TbSync account ID)
- `tbsyncSyncAccountByUser` (by user email, e.g. `jl4624@cornell.edu`)
- `listCalendars`
- `createCalendarEvent` (**v0.1.0: all-day only**)
- `syncAndCreateCalendarEvent` (pre-sync → create → post-sync; safest default)

## Documentation

Start here:
- `docs/USAGE.md` — quick-start + copy/paste curl examples
- `docs/TROUBLESHOOTING.md` — common issues (port not listening, hangs, dependency checks)
- `docs/PROCESS_RULES.md` — **operational rules / source of truth** for reliable TbSync+EAS workflows

## Development notes

This project is meant to be open-source friendly:
- minimal assumptions about the user’s Thunderbird setup
- defensive checks + explicit error messages when dependencies are missing
- careful handling of time zones and all-day events

## License

MIT. See `LICENSE`.

### Notices

This project includes/depends on components under other licenses (e.g., Mozilla/Thunderbird components under MPL-2.0). See `NOTICE.md`.
