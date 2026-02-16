# T4i New Hire Data Sync — Function Usage Guide (v8)

## Fully Automatic (No Action Needed)

| Function | Trigger | Description |
|----------|---------|-------------|
| `syncNewHireData()` | Runs automatically at 6 AM + 8 PM ET daily | Incremental sync of the last 6 weeks of tabs. You don't need to do anything — it runs on its own. Uses robust Source Tab matching to prevent duplicate accumulation. |

## Occasionally Useful (Run Manually When Needed)

| Function | When to Use | How to Run |
|----------|-------------|------------|
| `cleanupDuplicates()` | When you notice the same person appearing multiple times with the same Start Date and Source Tab | Select from function dropdown → Click Run → Check Log for removed count. Dedup key: First Name + Last Name + Start Date + Hire Type + Source Tab. |
| `discoverTabs()` | When you suspect source tabs were added/removed and want to confirm which tabs are in sync scope | Select from function dropdown → Click Run → Check Execution Log |
| `runSortAndFormat()` | After you manually edited data in the destination sheet and want to re-sort. Also re-enforces date and Source Tab column formats. | Select from function dropdown → Click Run |
| `cleanupIrregularDates()` | When you notice rows with "TBD", "N/A", or other non-date text in the Start Date or Ship Date columns. Also catches dates outside the 2000-2030 range. | Select from function dropdown → Click Run → Check Log to confirm how many rows were removed |
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

## v8 Data Integrity Notes

- **Source Tab column** is now enforced as plain text format. This prevents Google Sheets from auto-converting tab names like "February 16 2026" into date values, which previously caused duplicate rows to accumulate with every sync cycle.
- **Date columns** (Start Date, Ship Date) are normalized to midnight when synced. This eliminates hidden time components from source data that could cause inconsistent sorting.
- **All write operations** now use a shared `writeDestData_()` helper that enforces correct column formats every time, ensuring no format drift after cleanup or sync operations.
