# MPEC World Cup 2026 Picks Page

This folder contains a simple static picks page and a one-click JSON export pipeline.

## Files

- `index.html` — the public webpage
- `picks.json` — generated public picks data
- `export_picks_json.py` — reads the Excel workbook and writes `picks.json`
- `setup_once.bat` — installs the Python dependency `openpyxl`
- `update_picks_json.bat` — one-click exporter
- `preview_local_server.bat` — starts a local preview server

## First-time setup

1. Make sure Python is installed.
2. Double-click `setup_once.bat`.

## Update picks.json

Double-click:

```text
update_picks_json.bat
```

The script reads this workbook by default:

```text
C:\Users\kevin\OneDrive - Mass General Brigham\world_cup_2026_group_stage_challenge_MPEC_vCurrent.xlsx
```

and writes/overwrites:

```text
picks.json
```



## Privacy

The exporter includes only:

- name
- group rankings
- third-place qualifier groups
- tiebreaker goals

It intentionally excludes response IDs, timestamps, and raw Microsoft Forms metadata.
