#!/usr/bin/env python3
"""
diagnose.py — inspect what the dashboard actually sees in each file.

Run it the same way the app authenticates (env vars or token.json), or point
it at local files:

    python3 diagnose.py                          # OneDrive (uses .env / token.json)
    LOCAL_DATA_DIR=./sample_data python3 diagnose.py   # local folder

For every file it prints each sheet's columns and which ones were mapped to
User / Workingdate / Effort / Project Code, how many entries were read, and
any row/hour discrepancy against the source sheet.
"""

import sys

import server
from onedrive_source import DATA_EXTS, OneDriveSource


def tree():
    """List EVERY file under the folder, marking which ones the dashboard reads."""
    src = OneDriveSource()
    if src.local_dir:
        print("--tree only works against OneDrive (unset LOCAL_DATA_DIR).")
        return
    print(f"Folder: {src.folder!r}\n")
    items = src.walk(include_all=True)
    if not items:
        print("EMPTY — nothing found. Check ONEDRIVE_FOLDER spelling/casing.")
        return
    cur = None
    used = skipped = 0
    for it, rel, is_data in items:
        if rel != cur:
            cur = rel
            print(f"\n  {rel or '(base folder)'}/")
        size = it.get("size", 0)
        if is_data:
            used += 1
            print(f"      [READ]    {it['name']}  ({size:,} bytes)")
        else:
            skipped += 1
            print(f"      [ignored] {it['name']}  ({size:,} bytes)")
    print(f"\n{used} file(s) the dashboard reads, {skipped} ignored.")
    if skipped:
        print(f"Ignored files do not end in {', '.join(DATA_EXTS)} (or start with '~$').")


def main():
    if "--tree" in sys.argv:
        tree()
        return
    src = OneDriveSource()
    files = src.load_files()
    print(f"Source mode: {src.mode} · {len(files)} file(s)\n")

    all_warnings = []
    for name, year, raw in files:
        print("=" * 78)
        print(f"file        : {name}   (year folder {year!r})")
        print(f"sub team    : {server.team_from_filename(name)!r}")
        try:
            for label, score, df in server.read_tables(name, raw):
                cols = [str(c).strip() for c in df.columns]
                print(f"  sheet     : {label or '(csv)'}  ({len(df)} rows, header score {score})")
                print(f"  columns   : {cols}")
                m = server._map_columns(cols)
                if m is None:
                    print("  mapped    : -- skipped, no User/Workingdate/Effort columns --")
                    continue
                print(f"  mapped    : user={m['user']!r} date={m['date']!r} "
                      f"effort={m['effort']!r} project={m['proj']!r} task={m['task']!r}")
            warn = []
            entries = server.extract_file(name, raw, warn=warn)
            print(f"work entries: {len(entries)}  ({round(float(entries['eff'].sum()), 1)} hours)")
            if len(entries):
                counts = {}
                for d in entries["date"]:
                    y, m = server.period_of(d)
                    counts[f"{y}-{m:02d}"] = counts.get(f"{y}-{m:02d}", 0) + 1
                print("periods     : " + ", ".join(f"{k}({v})" for k, v in sorted(counts.items())))
                print(f"dates       : {min(entries['date'])} .. {max(entries['date'])}")

                # A dd/mm/yyyy column read as mm/dd/yyyy leaves a fingerprint:
                # every surviving day-of-month is <= 12 (the real month number),
                # and the rows with day > 12 vanished entirely.
                days = sorted({d.day for d in entries["date"]})
                months = {d.month for d in entries["date"]}
                if days and max(days) <= 12 and len(months) > 1:
                    print(f"  ** SUSPECT day/month swap: every day-of-month is <= 12 "
                          f"({days}) across {len(months)} months. If this file is "
                          f"dd/mm/yyyy, set DATE_DAYFIRST=1 and re-run.")
                print(f"days seen   : {days if len(days) <= 31 else '1..31'}")
            for w in warn:
                print(f"  ! {w}")
            all_warnings += warn
        except Exception as e:  # noqa: BLE001
            print(f"ERROR       : {type(e).__name__}: {e}")

    print("=" * 78)
    if all_warnings:
        print(f"\n{len(all_warnings)} warning(s) — these are where the dashboard and "
              f"the workbooks can disagree:")
        for w in all_warnings:
            print(f"  ! {w}")
    else:
        print("\nNo discrepancies: every row in every sheet was read.")
    print("\nNote: a dashboard month is the 26th->25th billing period, so "
          "'Jun 2025' = 26 May .. 25 Jun. It will not match a calendar-month "
          "total in Excel.")


if __name__ == "__main__":
    main()
