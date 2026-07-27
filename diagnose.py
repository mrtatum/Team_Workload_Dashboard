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

import server
from onedrive_source import OneDriveSource


def main():
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
                periods = sorted(set(entries["date"].map(server.period_of)))
                print(f"periods     : {[f'{y}-{m:02d}' for y, m in periods]}")
                print(f"dates       : {min(entries['date'])} .. {max(entries['date'])}")
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
