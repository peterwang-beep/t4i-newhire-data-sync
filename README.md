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
  Google Apps Script (v9)
  ├── fullSync()           → Batch initial load (all historical data)
  ├── syncNewHireData()    → Daily incremental sync (last 6 weeks)
  ├── cleanupDuplicates()  → Remove duplicate rows
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
| Production code | v9 |

## Project Files

| File | Description |
|------|-------------|
| `v9_Code.gs` | **Production code** (latest — validation rules, skip logging, improved dedup) |
| `v8_Code.gs` | Data integrity fixes (Source Tab format, dedup key) |
| `v7_Code.gs` | Added irregular date filter |
| `v6_Code.gs` | Performance fix (Range.sort) |
| `v1_Code.gs` – `v5_Code.gs` | Earlier version history |
| [`en/README.md`](en/README.md) | Full English documentation (architecture, column spec, function reference) |
| [`FUNCTION_GUIDE.md`](FUNCTION_GUIDE.md) | Function usage guide for daily operations |

## v9 Key Changes (Feb 2026)

| Change | Description |
|--------|--------------|
| **Two-part validation** | Part A: Start/Ship Date at least one non-empty; Part B: Last Name, Full name, or Username at least one non-empty |
| **Bad date format filter** | Skips rows with `7//22/2025`, `10/12.2021`, `3/3/3035` (double slash, dot separator, year out of range) |
| **Skip logging** | Per-sheet counts and examples for skipped rows (noDate, badFmt, noIdentity, tooOld, future) |
| **Improved dedup key** | `cleanupDuplicates` now includes Ship Date in key to avoid false deduplication |

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
