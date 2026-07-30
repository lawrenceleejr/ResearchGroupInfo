#!/usr/bin/env python3
"""Group-wide summary CSV generator.

Reads a directory of per-person Research Record spreadsheets (the same inputs
as ``generate_dashboard.py``) and writes summary CSVs aggregated across the
whole group:

    presentations.csv   every talk / poster / seminar, one row each
    publications.csv     every paper / note, one row each
    conferences.csv      unique conferences/meetings attended, with attendees
    schools.csv          unique schools / summer schools attended, with attendees
    grants.csv           the group's funded grant portfolio

Conferences and schools are drawn from two sources and merged: the dedicated
**Conferences & Schools** attendance tab, and the **Presentations** tab (each
presentation implies attendance at its venue). Entries whose type/venue looks
like a school / summer school / academy / tutorial are classified as schools;
everything else is a conference. Rows are grouped by event *series* (a trailing
year is stripped), so "USMCC 2025" and "USMCC 2026" collapse to one "USMCC" row
spanning both years, and the same person+event+year is counted once.

Usage:
    python generate_summaries.py <input_dir> [-o output_dir]
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import glob
import os
import re
import sys

from generate_dashboard import parse_workbook

STAMP = ""  # set in main(); appended to each CSV filename unless --no-datestamp


def out_path(out_dir: str, name: str) -> str:
    """Build an output path, inserting the date stamp before the extension."""
    if STAMP:
        root, ext = os.path.splitext(name)
        name = f"{root}_{STAMP}{ext}"
    return os.path.join(out_dir, name)

SCHOOL_RE = re.compile(r"\b(school|academy|tutorial)\b", re.I)
SCHOOL_TYPES = {"lecture"}


def _year(date: str) -> str:
    m = re.search(r"(?:19|20)\d{2}", date or "")
    return m.group(0) if m else ""


def _venue_series(venue: str) -> str:
    """Normalise a venue to its series name by stripping a trailing year."""
    v = (venue or "").strip()
    v = re.sub(r"[\s,'’]*(?:19|20)\d{2}\s*$", "", v).strip(" ,-–")
    return v or (venue or "").strip()


def _is_school(pres: dict) -> bool:
    if (pres.get("type") or "").strip().lower() in SCHOOL_TYPES:
        return True
    return bool(SCHOOL_RE.search(pres.get("venue") or ""))


def _sorted_join(values) -> str:
    return "; ".join(sorted(v for v in dict.fromkeys(values) if v))


def load_people(input_dir: str) -> list[dict]:
    paths = sorted(
        p for p in glob.glob(os.path.join(input_dir, "*.xlsx"))
        if not os.path.basename(p).startswith(("~$", "."))
        and "template" not in os.path.basename(p).lower()  # skip the blank template
    )
    if not paths:
        sys.exit(f"No .xlsx files found in {input_dir}")
    people = []
    for path in paths:
        try:
            people.append(parse_workbook(path))
        except Exception as exc:
            print(f"  SKIP {os.path.basename(path)}: {exc}", file=sys.stderr)
    if not people:
        sys.exit("No records could be parsed.")
    return people


def write_presentations(people, out_dir) -> int:
    path = out_path(out_dir, "presentations.csv")
    n = 0
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Person", "Type", "Venue / meeting", "Location", "Date",
                    "Title", "Invited?", "External travel funding", "Link"])
        for p in people:
            for r in p.get("presentations", []):
                w.writerow([p["name"], r.get("type", ""), r.get("venue", ""),
                            r.get("location", ""), r.get("date", ""), r.get("title", ""),
                            r.get("invited", ""), r.get("travel", ""), r.get("link", "")])
                n += 1
    print(f"  wrote {os.path.basename(path)}  ({n} rows)")
    return n


def write_publications(people, out_dir) -> int:
    path = out_path(out_dir, "publications.csv")
    n = 0
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Person", "Title", "Role", "Journal / collaboration",
                    "Status", "Date", "DOI / link"])
        for p in people:
            for r in p.get("publications", []):
                w.writerow([p["name"], r.get("title", ""), r.get("role", ""),
                            r.get("journal", ""), r.get("status", ""), r.get("date", ""),
                            r.get("doi", "")])
                n += 1
    print(f"  wrote {os.path.basename(path)}  ({n} rows)")
    return n


def _event_is_school(name: str, type_: str) -> bool:
    hay = f"{type_} {name}"
    return bool(SCHOOL_RE.search(hay))


def collect_attendance(people):
    """Unify attendance records from the Conferences & Schools tab and the
    Presentations tab into a flat list of dicts."""
    records = []
    for p in people:
        for r in p.get("events", []):
            name = (r.get("event") or "").strip()
            if not name:
                continue
            records.append({
                "person": p["name"], "name": name,
                "year": _year(r.get("start", "")),
                "location": r.get("location", ""),
                "descriptor": r.get("type", ""),
                "is_school": _event_is_school(name, r.get("type", "")),
            })
        for r in p.get("presentations", []):
            venue = (r.get("venue") or "").strip()
            if not venue:
                continue
            records.append({
                "person": p["name"], "name": venue,
                "year": _year(r.get("date", "")),
                "location": r.get("location", ""),
                "descriptor": r.get("type", ""),
                "is_school": _is_school(r),
            })
    return records


def _aggregate(records, want_schools: bool):
    """Group attendance records by event series, de-duplicating a repeated
    person+series+year."""
    agg: dict[str, dict] = {}
    for rec in records:
        if rec["is_school"] != want_schools:
            continue
        key = _venue_series(rec["name"]).lower()
        entry = agg.setdefault(key, {
            "name": _venue_series(rec["name"]), "attendees": [], "years": [],
            "locations": [], "types": [], "seen": set(), "count": 0,
        })
        dedup = (rec["person"].lower(), rec["year"])
        if dedup in entry["seen"]:
            continue
        entry["seen"].add(dedup)
        entry["attendees"].append(rec["person"])
        entry["years"].append(rec["year"])
        entry["locations"].append(rec["location"])
        entry["types"].append(rec["descriptor"])
        entry["count"] += 1
    return agg


def write_conferences(records, out_dir) -> int:
    path = out_path(out_dir, "conferences.csv")
    agg = _aggregate(records, want_schools=False)
    rows = sorted(agg.values(), key=lambda e: (-e["count"], e["name"].lower()))
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Conference / meeting", "Attendees", "# Attendees",
                    "Years", "Locations", "# Attendances", "Roles / contributions"])
        for e in rows:
            attendees = list(dict.fromkeys(e["attendees"]))
            w.writerow([e["name"], "; ".join(sorted(attendees)), len(attendees),
                        _sorted_join(e["years"]), _sorted_join(e["locations"]),
                        e["count"], _sorted_join(e["types"])])
    print(f"  wrote {os.path.basename(path)}  ({len(rows)} rows)")
    return len(rows)


def write_schools(records, out_dir) -> int:
    path = out_path(out_dir, "schools.csv")
    agg = _aggregate(records, want_schools=True)
    rows = sorted(agg.values(), key=lambda e: (-e["count"], e["name"].lower()))
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["School / course", "Attendees", "# Attendees",
                    "Years", "Locations", "# Attendances"])
        for e in rows:
            attendees = list(dict.fromkeys(e["attendees"]))
            w.writerow([e["name"], "; ".join(sorted(attendees)), len(attendees),
                        _sorted_join(e["years"]), _sorted_join(e["locations"]), e["count"]])
    print(f"  wrote {os.path.basename(path)}  ({len(rows)} rows)")
    return len(rows)


def write_grants(people, out_dir) -> int:
    path = out_path(out_dir, "grants.csv")
    n = 0
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Holder", "Grant / award title", "Agency / sponsor", "Role",
                    "Amount", "Start", "End", "Status", "Grant no. / notes"])
        for p in people:
            for r in p.get("grants", []):
                w.writerow([p["name"], r.get("title", ""), r.get("agency", ""),
                            r.get("role", ""), r.get("amount", ""), r.get("start", ""),
                            r.get("end", ""), r.get("status", ""), r.get("notes", "")])
                n += 1
    print(f"  wrote {os.path.basename(path)}  ({n} rows)")
    return n


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input_dir", help="directory containing the per-person .xlsx records")
    ap.add_argument("-o", "--output-dir", default=".", help="directory to write the CSVs into")
    ap.add_argument("--no-datestamp", action="store_true",
                    help="do not append the generation date to output filenames")
    args = ap.parse_args(argv)

    if not os.path.isdir(args.input_dir):
        sys.exit(f"Not a directory: {args.input_dir}")
    os.makedirs(args.output_dir, exist_ok=True)

    global STAMP
    STAMP = "" if args.no_datestamp else _dt.date.today().isoformat()

    people = load_people(args.input_dir)
    print(f"Parsed {len(people)} record(s). Writing summaries to {args.output_dir}/")
    write_presentations(people, args.output_dir)
    write_publications(people, args.output_dir)
    records = collect_attendance(people)
    write_conferences(records, args.output_dir)
    write_schools(records, args.output_dir)
    write_grants(people, args.output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
