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

2) **Use confirmed TbSync v5 operations for EAS**
   - Confirm the operation and supply the granted account and calendar IDs.
   - Event changes use the scoped calendar API so EAS records them.
   - Use a fresh event snapshot before an edit or deletion.

3) **Verify pre-sync and post-sync**
   - The write tools request one scoped refresh before the edit and one
     afterward, with provider checks and local readback.
   - Verify the result in Outlook before retrying an uncertain operation or
     making another write on a different machine.

4) **Don’t restart Thunderbird for syncing**
   - Restart only when needed to activate a newly-installed add-on build.

5) **Scheduled means “exists on Outlook calendar + has synced to cloud”**
   - If a task is marked scheduled in any task system, there must be a **real calendar event** created and visible in Outlook cloud.
   - When you notice “scheduled but not on calendar”, treat it as a **bug**: create the missing event immediately (or mark it unscheduled until fixed).

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

Check the account and folder in TbSync Manager. Once the calendar is bound,
select it again in the Bridge tab and inspect `tbsyncGetSyncStatus`.
Review unsynced local changes before using the manager's recovery controls.

## Morning planning pass (daily)

- The morning plan must be based on:
  - **Cornell calendar AND Duke calendar** (latest synced state)
  - `productivity/todo.csv` + `productivity/README.md`
- Always surface fixed-time commitments first (especially interviews).

## Operational Checklist (for risky sessions)

- Before making changes:
  - Confirm target calendar = Cornell (unless explicitly using Duke).
  - Confirm TbSync account user = `jl4624@cornell.edu`.

- After making changes:
  - Trigger **one** TbSync sync.
  - Wait and verify cloud on one device before touching the other.

## Tooling Notes

- Read through `listEvents` over a bounded date range.
- Keep the Bridge grant fixed during a request. To read Cornell and Duke,
  verify each account and select its calendar separately.
