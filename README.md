# Thunderbird TbSync MCP

MCP calendar tools for Thunderbird 153+, TbSync v5 and the Exchange
ActiveSync provider. The v5 extension checks the beta Bridge's capabilities
and account/calendar grant through the native helper. It uses the EAS
provider's existing change queue.

## Setup

Use TbSync 5.3.8 **beta** with the matching native helper (protocol 2) and
EAS-4-TbSync 5.3.11. In the TbSync Manager's Bridge tab, select one EAS
account and its calendar and leave unrestricted mode off. The normal
ATN build of TbSync does not include the Bridge tab. The helper's loopback
port is 47654. Its static header protects against ordinary cross-origin
browser requests. Other local processes can reach the helper while enabled.

Use the built `dist/thunderbird-tbsync-mcp.xpi` and point your MCP client
at this checkout's `mcp-bridge.cjs` with Node. The XPI starts on a local
port and publishes a per-session credential in
`${TMPDIR:-/tmp}/thunderbird-tbsync-mcp/connection.json` on Linux.
The stdio bridge reads that file automatically. If the client and Thunderbird
use different temp directories, set `TBSYNC_MCP_CONNECTION_FILE` to that
file's absolute path. The directory and file must be private to your user.
For manual writes, allow the MCP client up to four minutes per request.

`TBSYNC_MCP_BACKEND=v5` retains the earlier direct read-only mode. Use the
default `extension` mode for the authenticated XPI and the confirmed sync
and write workflows. Keep the Bridge grant fixed until a request finishes.

## Tools

- `tbsyncListAccounts` lists the single granted account and its account ID.
- `tbsyncGetSyncStatus` reports the granted calendar, last recorded sync,
  pending edits and queued calendar messages. A null count means the
  provider queue was unavailable.
- `listCalendars` and `listEvents` read the granted calendar. Event ranges
  need an explicit UTC offset or `Z` and are limited to 31 days.
  `listCalendarEvents` is an alias.
- `tbsyncSyncAccount` explicitly syncs the granted account, including its
  selected folders. It requires `accountID` and `confirm: true`. If TbSync
  reports an unchanged successful timestamp with no queued calendar work,
  the result includes `outcome: "unchanged"` and `migrationPending`. That
  confirms the observed state without claiming a new Exchange refresh.
- `createCalendarEvent`, `updateCalendarEvent`, and `deleteCalendarEvent`
  require `accountID`, `calendarId`, and explicit `confirm: true`. The
  workflow verifies one scoped refresh before the local edit and one
  afterward, then checks the provider queue and reads the event back.

The first write implementation handles individual events without meeting
participants or recurrence. For a create, use a unique stable `mcpTag`.
For edits and deletes, pass the event ID and the `ical` snapshot returned by
`listEvents` as `expectedIcal`. Read the event again if its snapshot changed.
A write can be applied locally even when the post-sync is uncertain.
Check the event and Outlook before retrying. Provider success and local
readback are different from an independent Outlook check.
See [usage](docs/USAGE.md) and [process rules](docs/PROCESS_RULES.md).

## Build and test

```sh
bash scripts/build.sh
node --test
```

To also test the iCalendar codec against Thunderbird's implementation,
set `TBSYNC_TEST_ICAL_MODULE` to an extracted
`modules/calendar/Ical.sys.mjs` from the installed application's
`omni.ja`. These tests use fixtures and temporary loopback servers.

The v4 extension and its legacy workflows are available on the original
branch. This package keeps the extension ID and requires Thunderbird 153.0
or later.

MIT license. See [notices](NOTICE.md).
