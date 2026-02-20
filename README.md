BIG THANKS for Cursor with claude-4.6-opus-max

# T4i New Hire Data Sync

Automated Google Apps Script pipeline that consolidates **T4i new hire onboarding records** from a 510-tab Google Spreadsheet into a single, clean, QlikSense-ready sheet — updated automatically twice daily.

## What It Does

- Scans **510 source tabs** in the T4i New Hire List, matches **448 tabs** with relevant columns
- Consolidates **30,000+ historical records** (Jan 2021 – present) into one unified sheet
- Runs **incremental daily sync** at 6:00 AM and 8:00 PM ET (only re-processes last 6 weeks)
- Filters out empty rows, future-dated entries, pre-2021 data, and irregular dates ("TBD", "N/A") automatically
- Prevents Source Tab column auto-formatting with enforced plain text format
- Normalizes all dates to midnight to eliminate hidden time component issues
- Tracks data origin with a "Source Tab" column for full traceability
- Output is formatted and ready for direct **QlikSense** loading

## Architecture

```
T4i New Hire List (510 tabs)
        │
        ▼
  Google Apps Script (v10)
  ├── fullSync()           → Batch initial load (all historical data)
  ├── syncNewHireData()    → Daily incremental sync (last 6 weeks)
  ├── cleanupDuplicates()  → Remove duplicate rows (SNOW Ticket in key, in-memory sort)
  ├── recall()             → Undo last sync/cleanup operation
  ├── discoverTabs()       → Scan & report tab structure
  └── setupDailyTriggers() → Auto-schedule (6am + 8pm ET)
        │
        ▼
  newHireAutoSyncData (single sheet, 30K+ rows, 19 columns)
        │
        ▼
  QlikSense / Analytics
```

## Quick Stats

| Metric | Value |
|--------|-------|
| Source tabs | 510 (448 matched) |
| Rows synced | 30,000+ |
| Date range | Jan 2021 – present |
| Auto-sync | Daily at 6 AM + 8 PM ET |
| Daily sync time | ~1-2 minutes |
| Production code | v10 |

## Project Files

| File | Description |
|------|-------------|
| `v10_Code.gs` | **Production code** (F1–F3 fixes, O1–O6 optimizations) |
| `v8_Optimized.gs` | Recall/Undo, dedup log, empty-date filter |
| `v9_Code.gs` | Validation rules, skip logging, improved dedup |
| `v8_Code.gs` | Data integrity fixes (Source Tab format, dedup key) |
| `v7_Code.gs` | Added irregular date filter |
| `v6_Code.gs` | Performance fix (Range.sort) |
| `v1_Code.gs` – `v5_Code.gs` | Earlier version history |
| [`en/README.md`](en/README.md) | Full English documentation (architecture, column spec, function reference) |
| [`FUNCTION_GUIDE.md`](FUNCTION_GUIDE.md) | Function usage guide for daily operations |

## v10 Key Changes (Feb 2026)

| Change | Description |
|--------|--------------|
| **F1: recentTabSet fix** | Incremental sync uses canonical Source Tab for matching — eliminates 700+ duplicates per sync. |
| **F2: Full-date canonical** | `getTabFullDate_()` extracts day; "February 16 2026" ≠ "February 23 2026" in dedup. |
| **F3: SNOW Ticket in dedup key** | Prevents false dedup when names are empty (e.g. RITM3936739 vs RITM3936578). |
| **O1–O6** | toLowerCase dedup; isEmpty outside loop; in-memory sort; saveUndoSnapshot_ returns handle; try-catch per tab; recall restores formats. |

**Note**: Rows with both Start Date and Ship Date empty are removed by `cleanupIrregularDates()` (manual), not skipped during sync.

## v8 Key Fixes (Feb 16, 2026)

| Bug | Severity | Fix |
|-----|----------|-----|
| Source Tab auto-formatting caused infinite row duplication | CRITICAL | Enforced plain text format + robust comparison with `normalizeSourceTab_()` |
| Date columns retained hidden time components | MODERATE | Normalized to midnight in `processSheet_()` |
| Cleanup functions lost column formatting after rewrite | MODERATE | Centralized writes via `writeDestData_()` |
| `hasIrregularDate_()` missed edge cases | MINOR | Added year range validation (2000-2030) |
| `sortAndFormatSheet_` did not protect Source Tab format | MINOR | Added `@` format enforcement after sort |
| No way to remove accumulated duplicates | NEW | Added `cleanupDuplicates()` function |

---

*Built by Peter Wang | T4i Care East | Started February 12, 2026*
