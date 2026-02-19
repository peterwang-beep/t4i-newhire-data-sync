# T4i New Hire Data Sync — Function Usage Guide (v9)

## Fully Automatic (No Action Needed)

| Function | Trigger | Description |
|----------|---------|-------------|
| `syncNewHireData()` | Runs automatically at 6 AM + 8 PM ET daily | Incremental sync of the last 6 weeks of tabs. You don't need to do anything — it runs on its own. Uses robust Source Tab matching to prevent duplicate accumulation. |

## Occasionally Useful (Run Manually When Needed)

| Function | When to Use | How to Run |
|----------|-------------|------------|
| `cleanupDuplicates()` | When you notice the same person appearing multiple times with the same Start Date, Ship Date, and Source Tab | Select from function dropdown → Click Run → Check Log for removed count. **v9 dedup key**: First Name + Last Name + Username + Start Date + **Ship Date** + Hire Type + Source Tab. |
| `discoverTabs()` | When you suspect source tabs were added/removed and want to confirm which tabs are in sync scope | Select from function dropdown → Click Run → Check Execution Log |
| `runSortAndFormat()` | After you manually edited data in the destination sheet and want to re-sort. Also re-enforces date and Source Tab column formats. | Select from function dropdown → Click Run |
| `cleanupIrregularDates()` | When you notice rows with "TBD", "N/A", or other non-date text in Start/Ship Date. **v9**: Also catches `7//22/2025`, `10/12.2021`, `3/3/3035` (double slash, dot separator, year out of 2000-2030). | Select from function dropdown → Click Run → Check Log to confirm how many rows were removed |
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
- **Data quality issues** — Run the appropriate cleanup function (`cleanupDuplicates`, `cleanupIrregularDates`, or `cleanupOldData`).
- **Want to check status** — Run `discoverTabs()` or visit the Executions page in Apps Script to review logs.
- **Disaster recovery** — Run `resetFullSync()` → `fullSync()` to rebuild all data from scratch.

## v9 Validation & Data Notes

- **Two-part row validation**: (A) Start Date and Ship Date must have at least one non-empty; (B) Last Name, Full name, or Username must have at least one non-empty. Rows failing either part are skipped.
- **Bad date format rejection**: Rows with `7//22/2025`, `10/12.2021`, `3/3/3035`, or similar malformed dates are skipped. Execution logs show counts and up to 10 examples per reason (noDate, badFmt, noIdentity, tooOld, future).
- **Improved dedup key**: `cleanupDuplicates` now includes Ship Date, so same person + same Start + same Ship + same tab = duplicate. Different Ship Dates are no longer collapsed.
- **Source Tab column** remains plain text format. **Date columns** are normalized to midnight. **All write operations** use `writeDestData_()` for consistent formatting.
