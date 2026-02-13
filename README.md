BIG THANKS for Cursor with claude-4.6-opus-max

# T4i New Hire Data Sync

Automated Google Apps Script pipeline that consolidates **T4i new hire onboarding records** from a 510-tab Google Spreadsheet into a single, clean, QlikSense-ready sheet — updated automatically twice daily.

## What It Does

- Scans **510 source tabs** in the T4i New Hire List, matches **448 tabs** with relevant columns
- Consolidates **29,865 historical records** (Jan 2021 – Feb 2026) into one unified sheet
- Runs **incremental daily sync** at 6:00 AM and 8:00 PM ET (only re-processes last 6 weeks)
- Filters out empty rows, future-dated entries, and pre-2021 data automatically
- Tracks data origin with a "Source Tab" column for full traceability
- Output is formatted and ready for direct **QlikSense** loading

## Architecture

```
T4i New Hire List (510 tabs)
        │
        ▼
  Google Apps Script
  ├── fullSync()         → Batch initial load (all historical data)
  ├── syncNewHireData()  → Daily incremental sync (last 6 weeks)
  ├── discoverTabs()     → Scan & report tab structure
  └── setupDailyTriggers() → Auto-schedule (6am + 8pm ET)
        │
        ▼
  newHireAutoSyncData (single sheet, 29,865 rows, 19 columns)
        │
        ▼
  QlikSense / Analytics
```

## Quick Stats

| Metric | Value |
|--------|-------|
| Source tabs | 510 (448 matched) |
| Rows synced | 30,000+ | (02/13/2026 6AM EST)
| Date range | Jan 2021 – Feb 2026 |
| Auto-sync | Daily at 6 AM + 8 PM ET |
| Daily sync time | ~1-2 minutes |
| Code versions | v1 through v6 |

## Project Files

| File | Description |
|------|-------------|
| `v6_Code.gs` | **Production code** (latest version) |
| `v1_Code.gs` – `v5_Code.gs` | Version history |
| [`en/README.md`](en/README.md) | Full English documentation (architecture, column spec, function reference) |


---

*Built by Peter Wang | T4i Care East | February 12, 2026*
