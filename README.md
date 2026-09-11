# LSE-CTUR Collaboration Tracker

A live-updating website tracking the country-year harmonisation and data-acquisition
progress of the MTUS+ project, split between the LSE and CTUR teams.

**Site:** https://qtnvqz3.github.io/lse-ctur-tracker/

## What this is

- `index.html` — the static site: an interactive world map (Plotly choropleth) showing
  each country's best current status, plus two searchable/filterable tables (**List 1 —
  Data In Hand** and **List 2 — Data Requests**) and an **Unclear** table for rows the
  source spreadsheets couldn't resolve confidently.
- `data/tracker.json` — the data behind the site. Regenerated periodically (see below)
  and read by the page at load time via `fetch()`; no rebuild step is needed after it
  changes.
- `scripts/generate_data.py` — reads the team's live tracking spreadsheets and rewrites
  `data/tracker.json`. Encodes the same classification rules (and the same manually
  curated cross-checks between conflicting spreadsheets) as the original
  `LSE_CTUR_Collaboration_Tracker.xlsx` handed to the team.
- `scripts/update_and_push.ps1` — wrapper that runs the generator and, if the data
  actually changed, commits and pushes it.

## Why the update runs locally, not as a GitHub Action

`generate_data.py` reads directly from the team's Dropbox and OneDrive-mounted
spreadsheets (`I:\MTUS+\...\0_management\*.xlsx` and
`I:\OneDrive\...\Juana Lamote's files - MTUS+\MTUS_tracker.xlsx`). Those paths only
exist on a machine that has those drives mounted — a GitHub Actions runner in the cloud
has no access to them. So instead of a scheduled GitHub Action, a **Windows Scheduled
Task on this machine** runs `update_and_push.ps1` every 12 hours; it pushes the
refreshed `data/tracker.json` straight to `main`, and because GitHub Pages serves that
file live (no build step), the site picks it up immediately.

If the spreadsheets ever move to a location reachable from the internet (e.g. a proper
database or an API), this can be switched to a real `schedule:`-triggered GitHub Action.

## Regenerating the data manually

```bash
py -3 scripts/generate_data.py
git add data/tracker.json
git commit -m "Update tracker data"
git push
```

## Status categories

**List 1 (have access):** Done · Being processed · Allocated, not started · Not allocated

**List 2 (no access yet):** Data received, not yet in shared folder · Dataset agreed, not
received · Request sent, still in discussion · Request sent, no response yet · Request
not sent yet

The map colours each country by the *best* status found across all of its rows in both
lists (e.g. a country that is "Done" for one survey year and has an open request for
another shows as "Done").

## Known data caveats

See the "Notes & Conflicts" section of `LSE_CTUR_Collaboration_Tracker_README.docx`
(in the team's Dropbox `0_management` folder) for a list of specific contradictions
found between the source spreadsheets (e.g. Norway 2010, Italy 2008, Basque Country
survey years) that were deliberately left un-resolved rather than guessed.
