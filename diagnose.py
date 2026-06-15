#!/usr/bin/env python3
"""
diagnose.py — inspect what the dashboard actually sees in each file.

Run it the same way the app authenticates (env vars or token.json), or point
it at local files:

    python3 diagnose.py                          # OneDrive (uses .env / token.json)
    LOCAL_DATA_DIR=./sample_data python3 diagnose.py   # local folder

For every file it prints the columns it detected and which ones it mapped to
User / Workingdate / Effort / Project Code, plus how many work entries it read.
"""

import server
from onedrive_source import OneDriveSource


def main():
    src = OneDriveSource()
    files = src.load_files()
    print(f"Source mode: {src.mode} · {len(files)} file(s)\n")
    for name, year, raw in files:
        print("=" * 70)
        print(f"file        : {name}   (year folder {year!r})")
        try:
            df = server.read_dataframe(name, raw)
            cols = [str(c).strip() for c in df.columns]
            print(f"columns     : {cols}")
            date_col = server._find_col(cols, ["workingdate", "workdate"]) or server._find_col(cols, ["date"])
            effort_col = server._find_col(cols, ["effort"])
            user_col = server._find_col(cols, ["username", "user", "employee", "staff", "resource", "name"])
            proj_col = server._find_col(cols, ["projectcode", "project"])
            print(f"mapped      : user={user_col!r} date={date_col!r} effort={effort_col!r} project={proj_col!r}")
            entries = server.extract_file(name, raw)
            print(f"work entries: {len(entries)}")
            if len(entries):
                periods = sorted(set(entries['date'].map(server.period_of)))
                print(f"periods     : {[f'{y}-{m:02d}' for y, m in periods]}")
        except Exception as e:  # noqa: BLE001
            print(f"ERROR       : {type(e).__name__}: {e}")
    print("=" * 70)


if __name__ == "__main__":
    main()
