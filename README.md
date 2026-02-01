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

## Planned tools

- `listCalendars` (read-only)
- `tbsyncListAccounts`
- `tbsyncSyncAccount` (by account ID)
- `tbsyncSyncAccountByUser` (match `jl4624@cornell.edu`, etc.)
- `createCalendarEvent` (all-day + timed events)
- `syncAndCreateCalendarEvent` (pre-sync -> create -> post-sync; safest default)

## Development notes

This project is meant to be open-source friendly:
- minimal assumptions about the user’s Thunderbird setup
- defensive checks + explicit error messages when dependencies are missing
- careful handling of time zones and all-day events

## License

TBD (recommend MIT to match the style of `thunderbird-mcp`).
