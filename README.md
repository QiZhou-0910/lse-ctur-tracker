# MTUS+ Project: UCL_LSE Coordination Tracker

A live-updating website tracking the country-year harmonisation and data-acquisition
progress of the MTUS+ project, split between the LSE and CTUR teams.

**Site:** https://qizhou-0910.github.io/lse-ctur-tracker/

## What this is

- `index.html` — the static site: an Overview tab with an interactive world map (Plotly
  choropleth) and a summary dashboard, plus separate tabs for **List 1 — Data In Hand**,
  **List 2 — Data Requests**, and **Unclear** (rows the source spreadsheets couldn't
  resolve confidently).
- `data/tracker.json` — the data behind the site. Regenerated periodically (see below)
  and read by the page at load time via `fetch()`; no rebuild step is needed after it
  changes.
- `scripts/generate_data.py` — reads the team's live tracking spreadsheets and rewrites
  `data/tracker.json`, applying the categorisation rules described below.
- `scripts/update_and_push.ps1` — wrapper that runs the generator and, if the data
  actually changed, commits and pushes it.

## Source spreadsheets

`generate_data.py` reads four spreadsheets maintained by the LSE and CTUR teams (not
included in this repo — only the derived, de-identified summary in `data/tracker.json`
is published here):

| File | Maintained by | Used for |
|---|---|---|
| `data_management.xlsx` | LSE/CTUR shared | The master tracker — one row per country-year survey, with a `phase`, `clean by`, `who contacted the NSO` and `notes about NSO contact` column. Primary source for both lists. |
| `data_acquisition.xlsx` | LSE/CTUR shared | Microdata acquisition status across the HETUS rounds and non-HETUS country-waves, with free-text notes on each contact's reply. Used to fill in detail `data_management.xlsx` doesn't have. |
| `nso_contact_tracking.xlsx` | LSE | Outreach log: mail dates and reply status for specific National Statistics Office contacts. |
| `MTUS_tracker.xlsx` ("Tracker" sheet) | CTUR | CTUR's own country-year harmonisation tracker — assignee and status, cross-checked against `data_management.xlsx`. |
| `MTUS_tracker.xlsx` ("Calendar" sheet) | CTUR | CTUR's weekly work-planning calendar, colour-coded by status (Done / Ongoing / Needs Peer Check / Interrupted). Used only as a source of individually team-confirmed status corrections for specific rows — see `CALENDAR_OVERRIDE` below. |

## Why the update runs locally, not as a GitHub Action

These spreadsheets live on the team's Dropbox and OneDrive, mounted as local network
drives on one team member's machine — a GitHub Actions runner in the cloud has no way to
reach them. So instead of a scheduled GitHub Action, a **Windows Scheduled Task on that
machine** runs `update_and_push.ps1` every 12 hours; it regenerates `data/tracker.json`
and, if anything changed, pushes it straight to `main`. Because GitHub Pages serves that
file live (no build step), the site picks up the change immediately.

If the spreadsheets ever move to a location reachable from the internet (e.g. a proper
database or an API), this can be switched to a real `schedule:`-triggered GitHub Action.

## Regenerating the data manually

```bash
py -3 scripts/generate_data.py
git add data/tracker.json
git commit -m "Update tracker data"
git push
```

## Categorisation rules

The full logic lives in `scripts/generate_data.py`; this section explains it so the
mapping from a spreadsheet cell to a category on the site is traceable.

### List 1 — Data In Hand

Every row starts from the `phase` column of `data_management.xlsx`:

| `phase` value | Category |
|---|---|
| `cleaned`, `cleaned/to check` | **Done** |
| `cleaning in progress/ctur`, `cleaning in progress`, `in progress` | **Being processed** |
| `data in dbox, not cleaned`, `data we have on server` | **Allocated, not started** if a team is recorded, otherwise **Not allocated** |

**Team attribution:** taken from the `clean by` column (LSE/CTUR) in
`data_management.xlsx` first. If that's blank, the row is cross-checked against CTUR's
`MTUS_tracker.xlsx` ("Tracker" sheet) for the same country and year — if that sheet
names an assignee, the row is counted as CTUR's.

**On the site**, this category + team pair is shown as two per-team icon columns
(CTUR / LSE) instead of one combined column:

- **Done** — for whichever team is responsible for a "Done" row
- **In progress** — for whichever team is responsible for a "Being processed" row
- **Assigned** — for whichever team is responsible for an "Allocated, not started" row
- **Not started yet** — for the other team, and for both teams on an unallocated row

If CTUR's Tracker sheet records a *different* status than `data_management.xlsx`'s
`phase` for the same row, that's shown as a "cross-check note" rather than silently
picking one file over the other.

**Calendar-tab corrections:** a small, explicitly named `CALENDAR_OVERRIDE` dictionary
in `scripts/generate_data.py` sets one or both team statuses directly for specific
rows the team has confirmed against the Calendar sheet's colour-coding — unlike the
Tracker-sheet cross-check above, this can mark *both* teams' progress independently on
one row (e.g. LSE already finished a survey that CTUR is now separately re-checking).
Only rows the team has actually confirmed are listed there; every other colour on the
Calendar tab is left alone.

### List 2 — Data Requests

Every row starts from the same `phase` column, restricted to the "no access yet"
values (`data we don't have`, `MTUS, we don't have microdata`,
`data we don't have, in Charmes (2026)`, `data we will have through Juana`), combined
with the `who contacted the NSO` and `notes about NSO contact` columns:

| Condition (checked in order) | Category |
|---|---|
| `phase` = `data we will have through Juana` | **Dataset agreed, not received** |
| No contact recorded, or notes read exactly "should be contacted" | **Request not sent yet** |
| Notes mention an email/contact being sent, with no reply recorded | **Request sent, no response yet** |
| Anything else (default fallback) | **Request not sent yet** |

**Manual cross-checks:** a specific set of country-years are then hand-corrected where
`data_acquisition.xlsx` or `nso_contact_tracking.xlsx` recorded a materially different,
more up-to-date status than `data_management.xlsx` did — for example, a contact saying
data had already been shared into the team's drop zone, or a mail-sent date that
`data_management.xlsx` didn't have on record. Each of these is listed individually in
the `ENRICH` dictionary in `scripts/generate_data.py`, together with the exact source
quote it's based on, so every override can be traced back to the original spreadsheet
rather than applied silently.

**Waves not in `data_management.xlsx` at all:** a handful of survey waves that only
appear in `data_acquisition.xlsx` — mainly the HETUS 2020 round for several European
countries, plus North Macedonia — are added directly, shown with a `Wave` (e.g.
`HETUS 2020`) instead of a `Year`, since `data_management.xlsx` has no row for them.

### Map colouring (Overview tab)

Each country is coloured by the single *best* status found across all of its rows in
either list — Done ranks best, Request not sent yet ranks worst (see the `RANK` ordering
in `scripts/generate_data.py`). A country that is "Done" for one survey year and has an
open request for another shows as "Done".

### Unclear

A few rows (Benin 1998; Vanuatu 1983, 1995, 1999) have no `phase` and no notes at all in
`data_management.xlsx` — not enough information to place them in either list, so they're
kept separate rather than guessed.

## Known data caveats

See the "Notes & Conflicts" section of the team's `LSE_CTUR_Collaboration_Tracker_README.docx`
for a list of specific contradictions found between the source spreadsheets (e.g. Norway
2010, Italy 2008, Basque Country survey years) that were deliberately left un-resolved
rather than guessed.
