# Thunderbird TbSync MCP — Process Rules

This document is the **single source of truth** for how we operate the Thunderbird + TbSync (EAS/Outlook) calendar workflow.

## Goals

- Make calendar operations (create/modify/delete) **reliable** and **cloud-consistent**.
- Avoid duplicates and prevent “split-brain” state across:
  - Ubuntu Thunderbird local cache
  - Outlook cloud
  - MacBook Thunderbird local cache

## Defaults

- **Default scheduling calendar:** `Jiachang Liu Cornell (Calendar)` unless explicitly asked to use Duke.
- **Time zone:** `America/New_York` for user-facing scheduling.
- **All-day events must be created as `allDay=true`**, not as 00:00–23:59 timed blocks.

## Golden Rules (read this first)

1) **Never edit TbSync changelog files manually**
   - Do not write to `~/.thunderbird/*/TbSync/changelog68.json` as an operational path.
   - It can fail to propagate and can destabilize folder/target selection.

2) **Use TbSync-native operations for EAS**
   Use the tools that call TbSync’s Lightning wrapper so TbSync records the correct `*_by_user` changelog entries:
   - Create: `tbsyncCreateEasEventNative` (sets `added_by_user`)
   - Delete: `tbsyncDeleteEasEventsNativeBySql` (sets `deleted_by_user`)
   - Modify/Move: `tbsyncModifyEasEventsBySql` (must ensure `modified_by_user` when applicable)

3) **Sync strategy: one sync, then wait**
   - Trigger **one** TbSync sync after a batch of changes.
   - Wait 2–5 minutes for Outlook propagation before doing more writes.
   - Do not “sync spam” on multiple machines at once (reduces cache divergence).

4) **Don’t restart Thunderbird for syncing**
   - Restart only when needed to activate a newly-installed add-on build.

## Idempotency and Tagging

- Every managed event should include a short tag at the **end of the title**, e.g.:
  - `Title (mcp:20260201-aimor)`
- Also store the same tag in `X-MCP-TAG`.
- Tags should be **unique per semantic event** (don’t reuse a tag for different items).

## Known Data/Display Gotchas

### All-day items and “floating” time

- All-day events may be stored as `event_start_tz='floating'` and at UTC midnight.
- SQL date-range queries using **local midnight→midnight America/New_York** can miss them.
- When verifying existence, prefer:
  - title/tag lookup, or
  - widened time windows, or
  - calendar UI verification.

### Timezone anomalies

- Some items can carry surprising tz fields (e.g., `America/Los_Angeles`) and display shifted times.

## Folder / Calendar Disappears (Ubuntu Repair Procedure)

Symptom:
- `Jiachang Liu Cornell (Calendar)` disappears from Ubuntu Thunderbird and from TbSync folder list.

Fix (programmatic):
1) Run `tbsyncResetEasFolderSync` for the Cornell account (forces FolderSync refresh).
2) Re-select the EAS folder `foldername="Calendar"` (usually `serverID="2"`) via `tbsyncSelectEasCalendarFolder`.
3) Confirm `listCalendars` includes `Jiachang Liu Cornell (Calendar)`.

## Operational Checklist (for risky sessions)

- Before making changes:
  - Confirm target calendar = Cornell.
  - Confirm TbSync account user = `jl4624@cornell.edu`.

- After making changes:
  - Trigger **one** TbSync sync.
  - Wait and verify cloud on one device before touching the other.

## Tooling Notes

- Provider enumeration (`getItems`, `getItem`) can hang; prefer SQL-based discovery + TbSync-native write operations.
- Avoid long-running blocking calls; prefer job-based flows when possible.
