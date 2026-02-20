# T4i New Hire Data Sync — Function Usage Guide (v10)

## Fully Automatic (No Action Needed)

| Function | Trigger | Description |
|----------|---------|-------------|
| `syncNewHireData()` | Runs automatically at 6 AM + 8 PM ET daily | Incremental sync of the last 6 weeks of tabs. You don't need to do anything — it runs on its own. Uses robust Source Tab matching to prevent duplicate accumulation. |

## Occasionally Useful (Run Manually When Needed)

| Function | When to Use | How to Run |
|----------|-------------|------------|
| `recall()` | **After you run any sync/cleanup by mistake** — restores the sheet to the state before that operation | T4i Sync menu → Recall (Undo last operation) → Check Log for confirmation |
| `cleanupDuplicates()` | When you notice the same person appearing multiple times with the same Start Date, Ship Date, and Source Tab | Select from function dropdown → Click Run → Log shows removed count and up to 30 example rows. **v10 dedup key**: First + Last + Username + **SNOW Ticket** + Start + Ship + HireType + Source Tab (canonical). |
| `discoverTabs()` | When you suspect source tabs were added/removed and want to confirm which tabs are in sync scope | Select from function dropdown → Click Run → Check Execution Log |
| `runSortAndFormat()` | After you manually edited data in the destination sheet and want to re-sort. Also re-enforces date and Source Tab column formats. | Select from function dropdown → Click Run |
| `cleanupIrregularDates()` | When you notice rows with "TBD", "N/A", or other non-date text in Start/Ship Date. Also removes rows where **both** Start and Ship Date are empty. Catches `7//22/2025`, `10/12.2021`, `3/3/3035`. | Select from function dropdown → Click Run → Check Log to confirm how many rows were removed |
| `cleanupOldData()` | When you find pre-2021 data that leaked into the destination sheet | Select from function dropdown → Click Run |

## Rarely Needed (Special Situations Only)

| Function | When to Use |
|----------|-------------|
| `resetFullSync()` then `fullSync()` | When you need to rebuild all data from scratch (e.g., changed EARLIEST_YEAR, or discovered a large-scale data issue). Run `resetFullSync` first, then `fullSync` — it will automatically process all tabs in batches. |
| `setupDailyTriggers()` | When triggers were deleted or need to be reconfigured. Under normal circumstances, this only needs to run once during initial setup. |
| `removeDailyTriggers()` | When you want to temporarily pause automatic syncing. |
| `removeTriggers()` | Emergency stop — removes ALL triggers (both fullSync continuation and daily sync). |

## Quick Reference

- **99% of the time** — Do nothing. The sync runs automatically.
- **Regret an operation?** — Run `recall()` to undo the last sync/cleanup (single-step undo).
- **Data quality issues** — Run the appropriate cleanup function (`cleanupDuplicates`, `cleanupIrregularDates`, or `cleanupOldData`).
- **Want to check status** — Run `discoverTabs()` or visit the Executions page in Apps Script to review logs.
- **Disaster recovery** — Run `resetFullSync()` → `fullSync()` to rebuild all data from scratch.

## v10 Notes

- **Sync**: Rows with both Start and Ship Date empty are **not** skipped during sync; use `cleanupIrregularDates()` to remove them.
- **Recall (Undo)**: Before each sync/cleanup, the current sheet is saved to a hidden `_Undo` sheet. Run `recall()` to restore. Only one level of undo.
- **Dedup key**: First + Last + Username + **SNOW Ticket** + Start + Ship + HireType + Source Tab (canonical). SNOW Ticket prevents false dedup when names are empty.
- **Empty-date filter**: `cleanupIrregularDates()` removes rows where both Start Date and Ship Date are empty.
