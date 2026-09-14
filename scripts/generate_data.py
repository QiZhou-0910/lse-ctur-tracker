# -*- coding: utf-8 -*-
"""
Regenerates data/tracker.json for the LSE-CTUR Collaboration Tracker site
from the live source spreadsheets. Meant to be run periodically (every 12h)
by a local scheduled task on a machine that has the Dropbox/OneDrive drives
mounted -- GitHub's own servers cannot see these paths, which is why this
does not run as a GitHub Action.

Sources:
  I:\\MTUS+\\Timeuseproject.dbox Dropbox\\Time Use Team Project\\1_CTUR_LSE_data\\0_management\\data_management.xlsx
  I:\\MTUS+\\...\\0_management\\data_acquisition.xlsx
  I:\\MTUS+\\...\\0_management\\data_acquisition\\nso_contact_tracking.xlsx
  I:\\OneDrive\\OneDrive - University College London\\Juana Lamote's files - MTUS+\\MTUS_tracker.xlsx
"""
import json
import re
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

BASE = Path(r"I:\MTUS+\Timeuseproject.dbox Dropbox\Time Use Team Project\1_CTUR_LSE_data\0_management")
ONEDRIVE = Path(r"I:\OneDrive\OneDrive - University College London\Juana Lamote's files - MTUS+")
OUT_PATH = Path(__file__).resolve().parent.parent / "data" / "tracker.json"


def norm(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def load_workbook_resilient(path, sheet=None):
    """OneDrive/Dropbox can hold a file lock; fall back to a temp copy."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / f"_lsectur_copy_{path.name}"
        shutil.copy2(path, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True)
    return wb[sheet] if sheet else wb


# ================= LOAD SOURCES =================
ws_t = load_workbook_resilient(ONEDRIVE / "MTUS_tracker.xlsx", "Tracker")
ctur_tracker = {}
for row in ws_t.iter_rows(min_row=2, values_only=True):
    country, year, harmtype, assigned, status, comments, roster, notcleaned = row
    if country is None:
        continue
    key = (norm(country), year)
    ctur_tracker.setdefault(key, []).append(dict(country=country, year=year, harmtype=harmtype,
                                                  assigned=assigned, status=status, comments=comments))

ws_m = load_workbook_resilient(BASE / "data_management.xlsx", "Feuil1")
mgmt_rows = []
for row in ws_m.iter_rows(min_row=2, values_only=True):
    _, country, code, year, survtype, phase, wholeaned, cleanby, whocontact, notes, why, source, _ = row
    if country is None:
        continue
    mgmt_rows.append(dict(country=country.strip() if isinstance(country, str) else country,
                           code=code, year=year, survtype=survtype, phase=phase,
                           wholeaned=wholeaned, cleanby=cleanby, whocontact=whocontact,
                           notes=notes, why=why, source=source))

ws_p = load_workbook_resilient(BASE / "data_cleaning" / "cleaning_priority.xlsx", "Sheet1")
priority_set = set()
for row in ws_p.iter_rows(min_row=3, values_only=True):
    _, country, code, year = row
    if country is None:
        continue
    priority_set.add((norm(country), year))

# ================= LIST 1 =================
HAVE_PHASES = {
    'cleaned': 'Done',
    'cleaned/to check': 'Done',
    'cleaning in progress/ctur': 'Being processed',
    'cleaning in progress': 'Being processed',
    'in progress': 'Being processed',
    'data in dbox, not cleaned': None,
    'data we have on server': None,
}
NOACCESS_PHASES = {"data we don't have", "MTUS, we don't have microdata",
                    "data we don't have, in Charmes (2026)"}
JUANA_PHASE = 'data we will have through Juana'

# Maps a row's overall category onto the four-state icon shown per team:
# whichever team is actually responsible gets 'done' / 'in_progress' /
# 'assigned'; the other team (and both, for unallocated rows) shows
# 'not_started'. This replaces a single "Team: LSE/CTUR" + "Status: ..."
# pair of columns with two per-team status columns, so the team a survey
# is assigned to and its progress are shown in one place instead of two.
CATEGORY_TO_STATUS = {
    'Done': 'done',
    'Being processed': 'in_progress',
    'Allocated, not started': 'assigned',
    'Not allocated': 'not_started',
}
STATUS_TO_CATEGORY = {v: k for k, v in CATEGORY_TO_STATUS.items()}
STATUS_RANK = {'done': 0, 'in_progress': 1, 'assigned': 2, 'not_started': 3}

# Cross-checked against CTUR's weekly work-planning "Calendar" tab in
# MTUS_tracker.xlsx (a second, independent source from the same file as the
# "Tracker" sheet above), confirmed with the team on 2026-09-11. Unlike the
# Tracker-sheet cross-check above, this can set BOTH teams' status on one
# row -- e.g. LSE already cleaned a survey that CTUR is now separately
# re-checking. Only entries the team actually confirmed are listed here;
# other ambiguous colours seen on the Calendar tab were left untouched.
CALENDAR_OVERRIDE = {
    (norm('Tanzania'), 2014): dict(
        ctur_status='done',
        note="Calendar tab (MTUS_tracker.xlsx): Tanzania 2014 confirmed done by the team."),
    (norm('Mongolia'), 2019): dict(
        ctur_status='in_progress',
        note="Calendar tab (MTUS_tracker.xlsx): CTUR (Sizhan) is currently re-checking this survey, independently of LSE's own completed pass."),
    (norm('United Kingdom'), 2014): dict(
        ctur_status='in_progress',
        note="Calendar tab (MTUS_tracker.xlsx): CTUR (Qi) is currently re-checking this survey, independently of LSE's own completed pass."),
    (norm('United Kingdom'), 2000): dict(
        ctur_status='in_progress',
        note="Calendar tab (MTUS_tracker.xlsx): CTUR (Qi) has started active work on this survey, ahead of the master tracker's 'not cleaned' phase."),
}

list1 = []
for r in mgmt_rows:
    phase = r['phase']
    if phase not in HAVE_PHASES:
        continue
    key = (norm(r['country']), r['year'])
    tmatches = ctur_tracker.get(key, [])
    category = HAVE_PHASES[phase]
    # Team (LSE/CTUR) is kept internally to derive the two status columns;
    # individual researcher names are deliberately not surfaced on the site.
    team = r['cleanby']
    ctur_tracker_status = tmatches[0]['status'] if tmatches else None
    if not team and tmatches and tmatches[0]['assigned']:
        team = 'CTUR'
    if category is None:
        category = 'Allocated, not started' if team else 'Not allocated'
    crosscheck = None
    if ctur_tracker_status and norm(ctur_tracker_status) not in norm(phase):
        crosscheck = f"CTUR Tracker shows status '{ctur_tracker_status}' -- verify against this row's phase."

    status = CATEGORY_TO_STATUS[category]
    if team == 'LSE':
        lse_status, ctur_status = status, 'not_started'
    elif team == 'CTUR':
        ctur_status, lse_status = status, 'not_started'
    else:
        ctur_status = lse_status = 'not_started'

    cal = CALENDAR_OVERRIDE.get((norm(r['country']), r['year']))
    if cal:
        if 'ctur_status' in cal:
            ctur_status = cal['ctur_status']
        if 'lse_status' in cal:
            lse_status = cal['lse_status']
        crosscheck = cal['note'] if not crosscheck else f"{crosscheck} | {cal['note']}"
        # The row's overall category (used by the map and the Overview
        # dashboard) follows whichever of the two team statuses is furthest
        # along, so a Calendar-tab correction is reflected there too.
        category = STATUS_TO_CATEGORY[min(ctur_status, lse_status, key=STATUS_RANK.get)]

    list1.append(dict(country=r['country'], code=r['code'], year=r['year'],
                       category=category, ctur_status=ctur_status, lse_status=lse_status,
                       source_phase=phase, source=r['source'], crosscheck_note=crosscheck))


def merge_duplicate_country_years(rows):
    """data_management.xlsx sometimes has two separate rows for the same
    country-year (one logged by each team, e.g. LSE marks it 'cleaned' while
    CTUR separately has it as 'cleaning in progress'). Collapse those into a
    single row per (country, year), taking each team's best (most advanced)
    status across its duplicate rows, so the site shows one row per survey."""
    groups = {}
    order = []
    for r in rows:
        key = (r['country'], r['year'])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)

    def uniq_join(values):
        seen = []
        for v in values:
            if v and v not in seen:
                seen.append(v)
        return '; '.join(seen) if seen else None

    merged = []
    for key in order:
        group = groups[key]
        if len(group) == 1:
            merged.append(group[0])
            continue
        ctur_status = min((g['ctur_status'] for g in group), key=STATUS_RANK.get)
        lse_status = min((g['lse_status'] for g in group), key=STATUS_RANK.get)
        category = STATUS_TO_CATEGORY[min(ctur_status, lse_status, key=STATUS_RANK.get)]
        merged.append(dict(
            country=group[0]['country'], code=group[0]['code'], year=group[0]['year'],
            category=category, ctur_status=ctur_status, lse_status=lse_status,
            source_phase=uniq_join(g['source_phase'] for g in group),
            source=uniq_join(g['source'] for g in group),
            crosscheck_note=uniq_join(g['crosscheck_note'] for g in group),
        ))
    return merged


list1 = merge_duplicate_country_years(list1)

# cleaning_priority.xlsx (data_cleaning/) flags specific country-years the
# team wants prioritised for cleaning. It only makes sense against surveys
# we actually hold, so it's applied to List 1 only; entries that don't match
# any List 1 row (e.g. because the survey turned out not to be acquired yet)
# are reported, not silently added.
matched_priority_keys = set()
for r in list1:
    key = (norm(r['country']), r['year'])
    r['priority'] = key in priority_set
    if r['priority']:
        matched_priority_keys.add(key)
unmatched_priority = sorted(priority_set - matched_priority_keys)
if unmatched_priority:
    print(f"WARNING: {len(unmatched_priority)} cleaning_priority.xlsx entries do not match any "
          f"List 1 row (survey not currently marked as in-hand): {unmatched_priority}")

# ================= LIST 2 =================
list2 = []


def add_l2(country, code, year, wave, category, team, note, origin):
    list2.append(dict(country=country, code=code, year=year, wave=wave, category=category,
                       team=team, note=note, origin=origin))


# Hand-curated cross-references between data_management.xlsx and the richer
# free-text notes in data_acquisition.xlsx / nso_contact_tracking.xlsx.
# See LSE_CTUR_Collaboration_Tracker_README.docx "Notes & Conflicts" for why
# each of these was resolved this way rather than left on the default category.
ENRICH = {
    (norm('Germany'), 2002): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: NSO 'just added to drop zone'; a further application for extended access is in progress."),
    (norm('Germany'), 2012): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: NSO 'just added to drop zone'; application for extended access in progress."),
    (norm('Germany'), 2022): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: CTUR's RA already has access to public microdata for 2020 and the previous two surveys; not yet reviewed."),
    (norm('Norway'), 2001): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: NSO contact said 'Just shared what we had' for 2000-2001."),
    (norm('Slovenia'), 2000): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: 'Just added to drop zone what we have' for 2000-2001."),
    (norm('Italy'), 2002): ('Data received, not yet in shared folder',
        "Per data_acquisition.xlsx: contact 'just shared a new folder called ITALY NEW' for 2002-2003."),
    (norm('Estonia'), 2021): ('Request sent, no response yet',
        "Per data_acquisition.xlsx: CTUR has an 'application in progress' for the HETUS 2020 wave."),
    (norm('Algeria'), 2012): (None,
        "Confirmed in nso_contact_tracking.xlsx: mail sent 19/05/2026, status 'awaiting response'."),
    (norm('Azerbaijan'), 2012): (None,
        "nso_contact_tracking.xlsx logs a mail sent 09/06/2026 for 'Azerbaijan' (wave unspecified), 'awaiting response'."),
    (norm('Azerbaijan'), 2008): (None,
        "nso_contact_tracking.xlsx logs a mail sent 09/06/2026 for 'Azerbaijan' (wave unspecified), 'awaiting response'."),
    (norm('Turkey'), 2015): (None,
        "nso_contact_tracking.xlsx logs a mail+form sent 19/05/2026 for 'Turkey' (wave unspecified), 'awaiting response'."),
    (norm('Turkey'), 2006): (None,
        "nso_contact_tracking.xlsx logs a mail+form sent 19/05/2026 for 'Turkey' (wave unspecified), 'awaiting response'."),
    (norm('Bulgaria'), 2010): (None,
        "FLAG: nso_contact_tracking.xlsx notes 'CHECK DROPZONE, I THINK I SAW BULGARIA' -- may already be received."),
    (norm('Belgium'), 2020): ('Request not sent yet',
        "Contingent on Belgium's NSO replying about the 1999/2005 roster issue; CTUR also plans to apply within weeks (data_acquisition.xlsx)."),
}

for r in mgmt_rows:
    phase = r['phase']
    if phase == JUANA_PHASE:
        cat, note = 'Dataset agreed, not received', None
    elif phase in NOACCESS_PHASES:
        key = (norm(r['country']), r['year'])
        if key in ENRICH:
            override_cat, extra_note = ENRICH[key]
            if not r['notes'] or r['notes'] == 'should be contacted':
                cat = override_cat or 'Request not sent yet'
            elif r['notes'] and ('sent' in norm(r['notes']) or 'email' in norm(r['notes'])):
                cat = override_cat or 'Request sent, no response yet'
            else:
                cat = override_cat
            note = extra_note
        else:
            if not r['notes'] and not r['whocontact']:
                cat = 'Request not sent yet'
            elif r['notes'] == 'should be contacted':
                cat = 'Request not sent yet'
            elif r['notes'] and ('sent' in norm(r['notes']) or 'emailed' in norm(r['notes'])):
                cat = 'Request sent, no response yet'
            else:
                cat = 'Request not sent yet'
            note = r['notes'] if r['notes'] not in (None, 'should be contacted') else None
    elif norm(r['country']) == norm('Belgium') and r['year'] == 2020 and phase is None:
        override_cat, extra_note = ENRICH[(norm('Belgium'), 2020)]
        cat, note = override_cat, (r['notes'] or '') + " | " + extra_note
    else:
        continue
    add_l2(r['country'], r['code'], r['year'], None, cat, r['whocontact'], note, 'data_management.xlsx')

NEW_WAVES = [
    ('North Macedonia', 'MKD', 'HETUS 2020', 'Request not sent yet', None,
     "Not in data_management.xlsx (which only lists 'Macedonia' 2009/2014); listed per data_acquisition.xlsx."),
    ('France', 'FRA', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Greece', 'GRC', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Italy', 'ITA', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Netherlands', 'NLD', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Poland', 'POL', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Romania', 'ROU', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Serbia', 'SRB', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Turkey', 'TUR', 'HETUS 2020', 'Request not sent yet', None, "Not in data_management.xlsx; listed per data_acquisition.xlsx."),
    ('Norway', 'NOR', 'HETUS 2020', 'Request sent, no response yet', None,
     "Per data_acquisition.xlsx: CTUR contact Elisabeth Ronning was approached ('yes, waiting for an answer')."),
]
for country, code, wave, cat, team, note in NEW_WAVES:
    add_l2(country, code, None, wave, cat, team, note, 'data_acquisition.xlsx (not in data_management.xlsx)')

UNCLEAR = [
    dict(country="Benin", code="BEN", year=1998,
         issue="No phase or notes recorded in data_management.xlsx. (Benin 2015 is a separate, already-cleaned survey.)"),
    dict(country="Vanuatu", code="VUT", year=1983, issue="No phase or notes recorded in data_management.xlsx."),
    dict(country="Vanuatu", code="VUT", year=1995, issue="No phase or notes recorded in data_management.xlsx."),
    dict(country="Vanuatu", code="VUT", year=1999, issue="No phase or notes recorded in data_management.xlsx."),
]

# ================= PER-COUNTRY MAP SUMMARY =================
RANK = {
    'Done': 1, 'Being processed': 2, 'Allocated, not started': 3, 'Not allocated': 4,
    'Data received, not yet in shared folder': 5, 'Dataset agreed, not received': 6,
    'Request sent, still in discussion': 7, 'Request sent, no response yet': 8,
    'Request not sent yet': 9,
}
country_best = {}
for r in list1 + list2:
    code = r.get('code')
    if not code:
        continue
    rank = RANK.get(r['category'], 99)
    cur = country_best.get(code)
    if cur is None or rank < cur['rank']:
        country_best[code] = dict(code=code, country=r['country'], category=r['category'], rank=rank)

map_data = sorted(country_best.values(), key=lambda x: x['country'] or '')

payload = dict(
    generated_at=datetime.now(timezone.utc).isoformat(),
    list1=sorted(list1, key=lambda x: (x['country'] or '', x['year'] if isinstance(x['year'], int) else 0)),
    list2=sorted(list2, key=lambda x: (x['country'] or '', x['year'] if isinstance(x['year'], int) else 9999)),
    unclear=UNCLEAR,
    map_data=map_data,
    category_rank=RANK,
)

OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
OUT_PATH.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
print(f"Wrote {OUT_PATH} -- List1: {len(list1)} rows, List2: {len(list2)} rows, "
      f"countries on map: {len(map_data)}")
