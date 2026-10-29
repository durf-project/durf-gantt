#!/usr/bin/env python3
"""Bake data/durf-roadmap.xlsx into index.html and export per-sheet CSVs.

    python3 tools/build.py

index.html keeps a block delimited by DURF-DATA markers. This script replaces the
JSON inside it with the workbook's raw sheet rows. The page normalises those rows
in JavaScript, using exactly the same code path as a workbook a user drops into
the page at runtime -- so a baked page and a freshly-loaded file can never drift.
"""
import csv
import difflib
import json
import os
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "durf-roadmap.xlsx"
HTML = ROOT / "gantt.html"
CSV_DIR = ROOT / "data" / "csv"

SKIP_SHEETS = {"READ ME"}
# Not a worksheet: carries the workbook's own metadata alongside the sheets, so
# baked data and a file loaded in the browser describe themselves the same way.
META_KEY = "__meta__"
# What the page cannot do without, per sheet. Headers are matched the way the
# page matches them: lower-cased and trimmed, nothing cleverer.
REQUIRED = {
    "Themes": ["theme_no", "name"],
    "Activities": ["theme_no", "activity", "start_month", "end_month"],
    "Settings": ["key", "value"],
}
# Everything the page reads. Any other column is yours to keep -- ignored, not
# an error -- which is how working columns such as "needs: ..." can live here.
KNOWN = {
    "Themes": {"theme_no", "name", "short_name", "colour", "color", "goal",
               "lead", "partners"},
    "Activities": {"theme_no", "activity", "outcomes", "tags", "start_month",
                   "end_month", "milestone_months", "links", "status",
                   "public_note", "public note", "note (public)", "notes"},
    "Budget": {"theme_no", "item", "description", "amount_eur", "shared"},
    "Open questions": {"theme_no", "question", "links"},
    "Settings": {"key", "value", "what it does"},
}
BEGIN = "/* == DURF-DATA:BEGIN == */"
END = "/* == DURF-DATA:END == */"


def cell(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "yes" if value else "no"
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, (int, float)):
        return value
    return str(value).strip()


def read_sheets(path):
    wb = load_workbook(path, data_only=True)
    modified = wb.properties.modified
    sheets = {META_KEY: {
        "modified": modified.replace(microsecond=0).isoformat() + "Z" if modified else "",
    }}
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        rows = []
        for raw in ws.iter_rows(values_only=True):
            row = [cell(v) for v in raw]
            while row and row[-1] == "":
                row.pop()
            rows.append(row)
        while rows and not any(str(v) for v in rows[-1]):
            rows.pop()
        if rows:
            sheets[ws.title] = rows
    return sheets


def say(kind, message):
    """Print a message; on GitHub Actions also raise it as an annotation, so it
    shows on the run summary rather than only inside the log."""
    if os.environ.get("GITHUB_ACTIONS"):
        print(f"::{kind}::{message.splitlines()[0]}")
    print(("ERROR: " if kind == "error" else "warning: ") + message)


def headers_of(rows):
    return [str(h).lower().strip() for h in (rows[0] if rows else []) if str(h).strip()]


def describe_missing(want, present):
    """Say what to look for, including the near-miss a human most likely typed."""
    squashed = {re.sub(r"[\s\-]+", "_", h): h for h in present}
    if want in squashed:
        return (f"'{want}' is written as '{squashed[want]}' -- the chart needs it "
                f"exactly as {want} (lower case, with an underscore)")
    close = difflib.get_close_matches(want, present, n=1, cutoff=0.7)
    if close:
        return f"'{want}' is missing -- is '{close[0]}' meant to be it?"
    return f"'{want}' is missing"


def validate(sheets):
    """Return a list of problems that must stop the build, and print the rest.

    A renamed or deleted column used to surface only as an empty or wrong chart:
    the page quietly drops rows it cannot place and defaults a missing month to
    1. Refusing here keeps the last good version published and names the cause.
    """
    errors = []
    for name, need in REQUIRED.items():
        if name not in sheets:
            errors.append(f"sheet '{name}' is missing "
                          f"(the workbook has: {', '.join(k for k in sheets if k != META_KEY)})")
            continue
        present = headers_of(sheets[name])
        for want in need:
            if want not in present:
                errors.append(f"sheet '{name}': " + describe_missing(want, present))

    for name, known in KNOWN.items():
        if name in sheets:
            extra = [h for h in headers_of(sheets[name]) if h not in known]
            if extra:
                print(f"  {name}: ignoring {len(extra)} column(s) the chart does not use: "
                      + "; ".join(extra))

    if errors:
        return errors

    # Problems that do not stop the build but make the chart quietly wrong.
    themes = indexed(sheets["Themes"])
    theme_nos = {str(r.get("theme_no")).strip() for r in themes}
    for i, r in enumerate(indexed(sheets["Activities"]), start=2):
        label = f"Activities row {i} ({str(r.get('activity'))[:40]})"
        if str(r.get("theme_no")).strip() not in theme_nos:
            say("warning", f"{label}: theme_no {r.get('theme_no')!r} is not in the "
                           "Themes sheet, so this activity will not appear")
        for col in ("start_month", "end_month"):
            if not isinstance(r.get(col), int):
                say("warning", f"{label}: {col} is {r.get(col)!r}, not a whole month "
                               "number -- the chart will treat it as month 1")
        if isinstance(r.get("start_month"), int) and isinstance(r.get("end_month"), int) \
                and r["end_month"] < r["start_month"]:
            say("warning", f"{label}: end_month {r['end_month']} is before "
                           f"start_month {r['start_month']}")
    return errors


def indexed(rows):
    head = [str(h).lower().strip() for h in rows[0]]
    out = []
    for row in rows[1:]:
        if not any(str(v).strip() for v in row):
            continue
        out.append({h: (row[i] if i < len(row) else "") for i, h in enumerate(head) if h})
    return out


def write_csvs(sheets):
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    for name, rows in sheets.items():
        if name == META_KEY:
            continue
        width = max(len(r) for r in rows)
        out = CSV_DIR / (name.lower().replace(" ", "-") + ".csv")
        with out.open("w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            for row in rows:
                w.writerow(list(row) + [""] * (width - len(row)))
        print(f"  {out.relative_to(ROOT)}")


def bake(sheets):
    html = HTML.read_text(encoding="utf-8")
    if BEGIN not in html or END not in html:
        sys.exit(f"marker block not found in {HTML.name}")
    payload = json.dumps(sheets, ensure_ascii=False, separators=(",", ":"))
    # </script> inside data would close the tag early.
    payload = payload.replace("</", "<\\/")
    block = f"{BEGIN}\nwindow.__DURF_SHEETS__ = {payload};\n{END}"
    html = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END), lambda _: block,
                  html, flags=re.S)
    HTML.write_text(html, encoding="utf-8")
    print(f"  {HTML.name}  ({len(payload) / 1024:.1f} kB of data)")


def main():
    if not XLSX.exists():
        sys.exit(f"missing {XLSX} -- run tools/seed_workbook.py first")
    sheets = read_sheets(XLSX)
    problems = validate(sheets)
    if problems:
        for problem in problems:
            say("error", problem)
        sys.exit("\nNot building: fix the columns above in the workbook. "
                 "The published chart keeps its last good version meanwhile.")
    print(f"read {XLSX.relative_to(ROOT)}: "
          + ", ".join(f"{k} ({len(v) - 1})" for k, v in sheets.items()
                      if k != META_KEY))
    print(f"  workbook last modified: {sheets[META_KEY]['modified'] or 'unknown'}")
    write_csvs(sheets)
    if HTML.exists():
        bake(sheets)
    else:
        print(f"  {HTML.name} not found yet -- skipped baking")


if __name__ == "__main__":
    main()
