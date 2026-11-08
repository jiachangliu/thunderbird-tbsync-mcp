# Thunderbird TbSync MCP v5 usage

Follow the [setup instructions](../README.md#setup), select the account and
calendar in the TbSync beta Bridge tab, then run the MCP client against
`mcp-bridge.cjs`.

Start with `tbsyncListAccounts`, `listCalendars`, and
`tbsyncGetSyncStatus`. Use `listEvents` with an explicit time range, for
example:

```json
{"start":"2026-11-08T00:00:00-05:00","end":"2026-11-09T00:00:00-05:00"}
```

For a confirmed write, use the account and calendar IDs from the read
tools. Here is a sample all-day create payload:

```json
{
  "accountID": "42",
  "calendarId": "granted-calendar-id",
  "confirm": true,
  "title": "Disposable sync test",
  "mcpTag": "disposable-sync-test",
  "allDay": true,
  "date": "2026-11-08"
}
```

All-day events accept an exclusive `endDate`. Timed events need
`allDay: false` and increasing `start` and `end` with an explicit offset
or `Z`. Create supports `title`, `description` and `location`. Update
patches those fields or the time range, and preserves other fields from
the current event. If updating an untagged event, supply a stable
`mcpTag` to make it a managed event.

Update and delete require `eventId` and `expectedIcal`. Use the exact
`ical` field returned for that item by `listEvents`, including its
line endings. Thunderbird's system timestamps are ignored when comparing
snapshots. Keep the grant fixed and avoid editing the same item in
Thunderbird while an MCP write is running.

For an initial live test, use one disposable event without attendees.
Confirm create, update and delete one at a time, and independently verify
each state in Outlook before proceeding. Compare `tbsyncGetSyncStatus`
before and after to check for pending changes or calendar messages.
