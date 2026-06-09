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

## Preview locally

Double-click:

```text
preview_local_server.bat
```

Then open:

```text
http://localhost:8000
```

## Share live

Host the folder as a static website. The page is just HTML + JSON.

Recommended options:

1. GitHub Pages
   - Put `index.html` and `picks.json` in a GitHub repository.
   - Enable GitHub Pages.
   - Each time you update `picks.json`, commit and push.

2. Netlify Drop
   - Drag this folder into Netlify Drop.
   - For updates, re-drag the folder after running `update_picks_json.bat`.

3. SharePoint/internal site
   - Best if you want MPEC-only access.
   - Upload/host the files in a location that serves `index.html` as a webpage.

## Privacy

The exporter includes only:

- name
- group rankings
- third-place qualifier groups
- tiebreaker goals

It intentionally excludes response IDs, emails, timestamps, and raw Microsoft Forms metadata.
