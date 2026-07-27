"""
server.py — Team Workload Dashboard backend (Flask)
==================================================
Pulls timesheet workbooks from OneDrive (Sysinfra_Workload, Excel), parses the
long/transactional schema, and serves both a JSON API and the HTML dashboard.

Schema (one row per work entry; a file may span many months):
    User Name | Workingdate | Effort | Project Code | ...

Business rules:
  - Rows are grouped into 26->25 billing periods by Workingdate; the dashboard
    shows one month per period (driven by the Year/Month dropdowns).
  - Effort is summed per person per day; a Project Code starting with
    NONBILL_PROJECT_PREFIX (default 'MFE') is non-billable, else billable.
  - Each period carries `workdays` (Mon-Fri minus holidays) for the 8h capacity.
  - Optional name exclusions via the EXCLUDE_NAMES env var (empty by default).

Endpoints:
    GET /            -> dashboard (static/index.html)
    GET /api/data    -> parsed JSON for all periods
    GET /api/health  -> {"status": "ok", "version": ...}
"""

from __future__ import annotations

import datetime
import io
import math
import os
import re
import sys
import time
import warnings

import pandas as pd
from flask import Flask, jsonify, request, send_from_directory

from onedrive_source import OneDriveError, OneDriveSource

app = Flask(__name__, static_folder="static")

VERSION = "2026-07-27-parsefixes"
CACHE_TTL = int(os.environ.get("CACHE_TTL_SECONDS", "300"))
STANDARD_DAY_HOURS = float(os.environ.get("STANDARD_DAY_HOURS", "8"))
HOLIDAYS_FILE = os.environ.get("HOLIDAYS_FILE", "holidays.txt")
_cache: dict = {"at": 0.0, "data": None}


def load_holidays():
    """
    Holiday dates to exclude from the normal-hours capacity. Sourced from:
      - env HOLIDAYS: comma/space separated ISO dates (2025-01-01 2025-04-14 ...)
      - a holidays file (HOLIDAYS_FILE, default holidays.txt): one ISO date per
        line; '#' starts a comment.
    Returns a set of datetime.date.
    """
    dates = set()

    def _add(tok):
        tok = tok.strip()
        if tok:
            try:
                dates.add(datetime.date.fromisoformat(tok))
            except ValueError:
                pass

    for tok in re.split(r"[,\s]+", os.environ.get("HOLIDAYS", "")):
        _add(tok)
    if os.path.exists(HOLIDAYS_FILE):
        with open(HOLIDAYS_FILE, encoding="utf-8") as fh:
            for line in fh:
                _add(line.split("#", 1)[0])
    return dates

META_PATTERNS = {
    "user": ("user name", "user", "name", "employee", "resource"),
    "total": ("total hours", "total", "hours"),
    "bill": ("% bill", "bill", "billable"),
    "nonbill": ("% non-bill", "non-bill", "nonbill", "non bill", "non billable"),
}

# Optional names to exclude from all data — comma-separated via the EXCLUDE_NAMES
# env var. Empty by default (no one excluded).
EXCLUDE_NAMES = {n.strip().lower() for n in os.environ.get("EXCLUDE_NAMES", "").split(",") if n.strip()}

# Sub team derived from the cost-center code in the FILENAME (e.g. *_HI05010.xlsx).
SUBTEAM_MAP = {"HI05010": "Presales", "HI05020": "Engineers"}
SUBTEAM_ORDER = ["Presales", "Engineers"]


def team_from_filename(name):
    m = re.search(r"(HI\d{5})", str(name), re.IGNORECASE)
    if m:
        code = m.group(1).upper()
        return SUBTEAM_MAP.get(code, code)
    return None
# Columns that are NOT day columns and must never be summed as daily hours.
IGNORE_DAY_COLS = {"sum", "total", "grand total", "total hours", "subtotal"}

MONTHS = {m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], start=1)}


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace("%", "").replace(",", "")
    if s == "" or s.lower() in ("nan", "na", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _pct(v):
    n = _num(v)
    if n is None:
        return None
    return round(n * 100, 1) if n <= 1 else round(n, 1)


def _match_meta(columns):
    lower = {c: str(c).strip().lower() for c in columns}
    found = {}
    for key, pats in META_PATTERNS.items():
        for c in columns:
            if c in found.values():
                continue
            if any(lower[c] == p or p in lower[c] for p in pats):
                found[key] = c
                break
    return found


def _month_sortkey(name):
    stem = re.sub(r"\.csv$", "", name, flags=re.IGNORECASE)
    m = re.search(r"(20\d{2})[-_ ]?(\d{1,2})", stem)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m = re.search(r"([A-Za-z]+)[-_ ]?(20\d{2})", stem)
    if m and m.group(1).lower() in MONTHS:
        return (int(m.group(2)), MONTHS[m.group(1).lower()])
    return (9999, 99)


MONTH_ABBR = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _month_from_name(name):
    """Extract the month (1-12) from the filename."""
    stem = re.sub(r"\.csv$", "", name, flags=re.IGNORECASE).strip()
    low = stem.lower()
    # month name (full or 3-letter), tolerant of underscores/spaces around it
    for mname, idx in MONTHS.items():
        if re.search(r"(?<![a-z])" + mname[:3], low):
            return idx
    m = re.search(r"(20\d{2})[-_ ]?(\d{1,2})", stem)        # e.g. 2025-06
    if m:
        return int(m.group(2))
    m = re.search(r"(?<!\d)(0?[1-9]|1[0-2])(?!\d)", stem)   # bare month number
    if m:
        return int(m.group(1))
    return None


def _year_from_folder(folder):
    """Extract the 4-digit year from the folder name."""
    if not folder:
        return None
    m = re.search(r"(20\d{2})", str(folder))
    return int(m.group(1)) if m else None


def _to_int_day(label):
    try:
        return int(float(str(label).strip()))
    except (ValueError, TypeError):
        return None


def _map_dates(day_cols, year, month):
    """
    Map each day-of-month column to a real date. The 26->25 window spans two
    months: high day numbers (26..end) belong to the previous month, and the
    sequence wraps (e.g. 31 -> 1) into `month`. Returns list of date|None.
    """
    if not year or not month:
        return [None] * len(day_cols)
    pm, py = (month - 1, year) if month > 1 else (12, year - 1)
    cur_m, cur_y = pm, py
    out, prev = [], None
    for c in day_cols:
        # full date in the header?
        try:
            dt = pd.to_datetime(c, errors="raise")
            out.append(dt.date())
            prev = dt.day
            continue
        except Exception:
            pass
        n = _to_int_day(c)
        if n is None:
            out.append(None)
            continue
        if prev is not None and n < prev:     # wrapped into the current month
            cur_m, cur_y = month, year
        try:
            out.append(datetime.date(cur_y, cur_m, n))
        except ValueError:
            out.append(None)
        prev = n
    return out


def _period_dates(year, month):
    """Full 26->25 calendar for the period, regardless of the CSV's columns."""
    pm, py = (month - 1, year) if month > 1 else (12, year - 1)
    start = datetime.date(py, pm, 26)
    end = datetime.date(year, month, 25)
    out, d = [], start
    while d <= end:
        out.append(d)
        d += datetime.timedelta(days=1)
    return out


NONBILL_PREFIX = os.environ.get("NONBILL_PROJECT_PREFIX", "MFE").strip().upper()


_EXACT, _PREFIX, _CONTAINS = 0, 1, 2


def _norm_col(c):
    """Column name reduced to letters+digits: 'Project_Code ' -> 'projectcode'."""
    return re.sub(r"[^a-z0-9]", "", str(c).strip().lower())


def _find_col(cols, patterns, exclude=()):
    """
    Pick the best column for a role. `patterns` are in priority order.

    Ranking is (pattern priority, match quality, position) — an exact match
    beats a prefix match beats a substring match, and only then does the
    left-most column win. The old version returned the first column that merely
    *contained* any keyword, so a sheet with 'Customer Name' or 'Project Name'
    to the left of 'User Name' / 'Project Code' silently mapped people onto
    customers and made every project look billable.
    """
    best, best_key = None, None
    for pos, c in enumerate(cols):
        if c in exclude:
            continue
        n = _norm_col(c)
        if not n:
            continue
        for pi, p in enumerate(patterns):
            p = _norm_col(p)
            if not p:
                continue
            if n == p:
                q = _EXACT
            elif n.startswith(p):
                q = _PREFIX
            elif p in n:
                q = _CONTAINS
            else:
                continue
            key = (pi, q, pos)
            if best_key is None or key < best_key:
                best, best_key = c, key
            break                      # highest-priority pattern for this column
    return best


def _valid_name(v):
    return v is not None and str(v).strip() != "" and str(v).strip().lower() != "nan"


# DATE_DAYFIRST: "1"/"0" forces day-first / month-first parsing. Unset = auto.
_DAYFIRST_ENV = os.environ.get("DATE_DAYFIRST", "").strip().lower()
DATE_DAYFIRST = {"1": True, "true": True, "yes": True,
                 "0": False, "false": False, "no": False}.get(_DAYFIRST_ENV)


def _parse_dates_series(series, warn=None, where=""):
    """
    Parse a date column, picking the interpretation that loses fewer rows.

    The old rule only retried day-first when MORE than half the rows failed, so
    a dd/mm/yyyy export sat right on the boundary: rows with day > 12 became
    NaT and were dropped, and the rest (e.g. 02/06 -> 6 Feb) landed in the wrong
    month. Set DATE_DAYFIRST to override the automatic choice.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")   # real Excel dates: no ambiguity

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        mfirst = pd.to_datetime(series, errors="coerce")
        dfirst = pd.to_datetime(series, errors="coerce", dayfirst=True)

    if DATE_DAYFIRST is not None:
        return dfirst if DATE_DAYFIRST else mfirst

    n_m, n_d = int(mfirst.notna().sum()), int(dfirst.notna().sum())
    if n_d != n_m:
        return dfirst if n_d > n_m else mfirst

    # Both parse the same number of rows. Year-first values (2025-06-02) are
    # unambiguous whatever pandas' dayfirst inference does with them.
    txt = series.astype(str).str.strip()
    txt = txt[txt.str.lower().isin(("", "nan", "none")) == False]  # noqa: E712
    if len(txt) and txt.str.match(r"^\d{4}[-/.]").mean() > 0.5:
        return mfirst

    # Otherwise, if the two readings disagree the column is genuinely ambiguous
    # (every day/month part <= 12) — keep the historical month-first reading but
    # say so.
    if warn is not None and n_m and not mfirst.equals(dfirst):
        warn.append(
            f"{where}: date column is ambiguous (every value fits both dd/mm and "
            f"mm/dd). Read as month-first; set DATE_DAYFIRST=1 if it should be "
            f"day-first.")
    return mfirst


def period_of(d):
    """The 26->25 billing period (year, month) that a date belongs to.
    Days 26..end fall into the NEXT month's period."""
    if d.day >= 26:
        m, y = (d.month + 1, d.year)
        if m > 12:
            m, y = 1, d.year + 1
        return (y, m)
    return (d.year, d.month)


def _nonblank(series):
    s = series.astype(str).str.strip().str.lower()
    return series.notna() & (s != "") & (s != "nan") & (s != "none")


def extract_long_entries(df, user_col, date_col, effort_col, proj_col, task_col=None,
                         cust_col=None, warn=None, where=""):
    """
    Flatten a long/transactional sheet into rows: (name, date, eff, billable,
    task, proj, cust). Vectorized + coercion-safe: bad effort -> 0, bad date ->
    NaT (dropped). A project code starting with NONBILL_PREFIX (default 'MFE')
    is non-billable.

    Coercion is no longer silent: anything dropped or zeroed is appended to
    `warn` so the totals can be reconciled against the source workbook.
    """
    df = df.reset_index(drop=True)

    raw_eff = df[effort_col]
    effort = pd.to_numeric(raw_eff, errors="coerce")
    bad_eff = int((effort.isna() & _nonblank(raw_eff)).sum())
    effort = effort.fillna(0.0)

    raw_dates = df[date_col]
    parsed = _parse_dates_series(raw_dates, warn=warn, where=where)
    bad_dates = int((parsed.isna() & _nonblank(raw_dates)).sum())
    dates = parsed.dt.date

    names = df[user_col].astype(str).str.strip()
    if proj_col is not None:
        proj = df[proj_col].astype(str).str.strip().str.upper()
        billable = ~proj.str.startswith(NONBILL_PREFIX)
    else:
        proj = pd.Series([""] * len(df))
        billable = pd.Series(True, index=df.index)
    if task_col is not None:
        task = df[task_col].astype(str).str.strip().replace({"": "Unspecified", "nan": "Unspecified"})
    else:
        task = pd.Series(["Unspecified"] * len(df))
    if cust_col is not None:
        cust = df[cust_col].astype(str).str.strip().replace({"nan": ""})
    else:
        cust = pd.Series([""] * len(df))

    work = pd.DataFrame({"name": names, "date": dates, "eff": effort,
                         "bill": billable.values, "task": task.values,
                         "proj": proj.values, "cust": cust.values})
    hours_in = float(effort.sum())
    rows_in = len(work)

    nm_l = work["name"].str.lower()
    keep_name = (work["name"] != "") & (nm_l != "nan")
    excluded = int((keep_name & nm_l.isin(EXCLUDE_NAMES)).sum())
    bad_names = int((~keep_name).sum())
    work = work[keep_name & (~nm_l.isin(EXCLUDE_NAMES))]
    work = work[work["date"].notna()]

    if warn is not None:
        hours_out = float(work["eff"].sum())
        lost = round(hours_in - hours_out, 2)
        bits = []
        if bad_dates:
            bits.append(f"{bad_dates} row(s) with an unreadable date")
        if bad_names:
            bits.append(f"{bad_names} row(s) with no user name")
        if excluded:
            bits.append(f"{excluded} row(s) excluded by EXCLUDE_NAMES")
        if bad_eff:
            bits.append(f"{bad_eff} non-numeric Effort cell(s) counted as 0")
        if bits or abs(lost) > 0.05:
            warn.append(
                f"{where}: kept {len(work)}/{rows_in} rows, "
                f"{round(hours_out, 1)}/{round(hours_in, 1)} hours"
                + (f" ({lost:+.1f}h vs the sheet)" if abs(lost) > 0.05 else "")
                + (" — " + "; ".join(bits) if bits else ""))
    return work


# Map a raw "Job task" value to a category by keyword. Order matters — the first
# matching keyword wins, so more specific codes (presale, pma, cma) come before
# the shorter ones they contain (ma).
TASK_RULES = [
    ("dev,imp,ma,wa", "Presales"),         # "Dev,Imp,MA,WA" presales variants
    ("presale", "Presales"),
    ("meeting", "Meeting"),                # incl. "Meeting/ Admin/ Support-MFE-HI05000*"
    ("leave", "Leave"),
    ("standby", "Standby"),
    ("training", "Training"),
    ("pma", "Preventive Maintenance"),
    ("cma", "Corrective Maintenance"),
    ("imp", "Implement"),
    ("ma", "Corrective Maintenance"),      # bare *MA* falls back to corrective
]


def map_task(raw):
    s = str(raw).strip().lower()
    if not s or s == "nan":
        return "Others"
    s = s.replace(" ", "")                 # tolerate spacing variants
    for kw, cat in TASK_RULES:
        if kw in s:
            return cat
    return "Others"


def build_period_month(year, month, work, holidays):
    """Build one dashboard month (26->25 period) from its entries."""
    canon = _period_dates(year, month)
    idx = {d: i for i, d in enumerate(canon)}
    days_out = [str(d.day) for d in canon]
    dates_out = [d.isoformat() for d in canon]
    workdays = sum(1 for d in canon if d.weekday() < 5 and d not in holidays)
    holiday_dates = [d.isoformat() for d in canon if d.weekday() < 5 and d in holidays]

    people = []
    for nm, g in work.groupby("name", sort=True):
        total = float(g["eff"].sum())
        billh = float(g.loc[g["bill"], "eff"].sum())
        nonh = float(g.loc[~g["bill"], "eff"].sum())
        daily = [0.0] * len(canon)
        for d, e in zip(g["date"], g["eff"]):
            if d in idx:
                daily[idx[d]] += float(e)
        team = None
        if "team" in g.columns:
            team = next((t for t in g["team"] if t and str(t).lower() != "nan"), None)
        tasks = {}
        standby_nb = 0.0
        if "task" in g.columns:
            for t, e, b in zip(g["task"], g["eff"], g["bill"]):
                key = map_task(t)                      # group raw job task -> category
                tasks[key] = tasks.get(key, 0.0) + float(e)
                if key == "Standby" and not b:         # standby on MFE* project
                    standby_nb += float(e)
            tasks = {k: round(v, 1) for k, v in tasks.items()}
        people.append({
            "name": nm,
            "team": team,
            "total": round(total, 1),
            "bill": round(billh / total * 100, 1) if total > 0 else None,
            "nonbill": round(nonh / total * 100, 1) if total > 0 else None,
            "daily": [round(x, 2) for x in daily],
            "tasks": tasks,
            "standbyNB": round(standby_nb, 1),   # Standby hours on MFE* projects
        })

    return {
        "label": f"{MONTH_ABBR[month]} {year}", "year": year, "month": month,
        "days": days_out, "dates": dates_out, "workdays": workdays,
        "holidayDates": holiday_dates,
        "normalHours": round(workdays * STANDARD_DAY_HOURS, 1),
        "standardDayHours": STANDARD_DAY_HOURS, "people": people,
    }


HEADER_TOKENS = (
    "workingdate", "workdate", "effort", "projectcode", "project",
    "username", "user", "employee", "staff", "resource", "fullname",
    "costcenter", "costcentre", "totalhours", "%bill", "nonbill", "date",
)


def _row_score(values):
    s = 0
    for v in values:
        t = str(v).strip().lower().replace(" ", "")
        if t in ("", "nan", "none"):
            continue
        if any(tok in t for tok in HEADER_TOKENS):
            s += 1
    return s


def _best_header_row(rdf):
    """Within a header=None DataFrame, find the row that looks most like a header."""
    best_i, best_s = 0, -1
    for i in range(min(len(rdf), 15)):
        s = _row_score(list(rdf.iloc[i].values))
        if s > best_s:
            best_i, best_s = i, s
    return best_i, best_s


def _dedupe(names):
    seen, out = {}, []
    for n in names:
        n = n if (n and n.lower() not in ("nan", "none")) else "col"
        if n in seen:
            seen[n] += 1
            out.append(f"{n}.{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out


def _read_excel_sheets(raw, low):
    """
    Read every sheet (header=None). Prefer the 'calamine' engine: it tolerates
    cells that the file mislabels as numeric but that actually hold text (a
    common exporter quirk that makes openpyxl raise
    'could not convert string to float'). Falls back to openpyxl/default.
    """
    engines = ["calamine", "openpyxl"] if low.endswith((".xlsx", ".xlsm")) else ["calamine", None]
    last = None
    for eng in engines:
        try:
            return pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None, engine=eng)
        except Exception as e:  # noqa: BLE001
            last = e
    raise last


def _table_from(rdf):
    """A header=None sheet -> (header score, body frame with real column names)."""
    hi, hs = _best_header_row(rdf)
    header = _dedupe([str(x).strip() for x in rdf.iloc[hi].tolist()])
    body = rdf.iloc[hi + 1:].copy()
    body.columns = header
    return hs, body.dropna(how="all").reset_index(drop=True)


def read_tables(name, raw):
    """
    EVERY table in the file as [(sheet label, score, body df), ...], best score
    first. For each sheet the header row is located (tolerating title/blank rows
    above the table) and duplicate column names are suffixed.

    This replaces the old "pick the single best-matching sheet" behaviour, which
    silently discarded data whenever a workbook split its timesheet across tabs
    (e.g. one per sub team).
    """
    low = name.lower()
    if low.endswith((".xlsx", ".xlsm", ".xls")):
        sheets = _read_excel_sheets(raw, low)
    else:
        sheets = {"": pd.read_csv(io.BytesIO(raw), header=None, dtype=object)}

    out = []
    for label, rdf in sheets.items():
        if rdf is None or rdf.empty:
            continue
        score, body = _table_from(rdf)
        if not body.empty:
            out.append((str(label), score, body))
    if not out:
        raise ValueError("workbook has no readable sheet")
    out.sort(key=lambda t: -t[1])
    return out


def read_dataframe(name, raw):
    """The single best-matching table in the file (used by diagnose.py)."""
    return read_tables(name, raw)[0][2]


def _map_columns(cols):
    """Map sheet columns onto roles, or None if the required three are missing."""
    date_col = _find_col(cols, ["workingdate", "workdate"]) or _find_col(cols, ["date"])
    effort_col = _find_col(cols, ["effort"])
    used = [c for c in (date_col, effort_col) if c]
    user_col = _find_col(cols, ["username", "user", "employee", "staff", "resource", "name"],
                         exclude=used)
    if not (date_col and effort_col and user_col):
        return None
    used.append(user_col)
    proj_col = _find_col(cols, ["projectcode", "project"], exclude=used)
    if proj_col:
        used.append(proj_col)
    task_col = _find_col(cols, ["jobtask", "tasktype", "task"], exclude=used)
    if task_col:
        used.append(task_col)
    cust_col = _find_col(cols, ["customer", "custname", "client"], exclude=used)
    return {"date": date_col, "effort": effort_col, "user": user_col,
            "proj": proj_col, "task": task_col, "cust": cust_col}


def extract_file(name, raw, warn=None):
    """
    Read a file and return its long-format work entries (name, date, eff, bill),
    combining EVERY sheet that carries the long/transactional schema. Sheets
    that don't (README, pivots, notes) are skipped; byte-identical duplicate
    sheets are counted once.
    """
    frames, seen, sigs, mapped_any = [], [], set(), False
    for label, _score, df in read_tables(name, raw):
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        cols = list(df.columns)
        seen.append((label, cols))
        m = _map_columns(cols)
        if m is None:
            continue
        mapped_any = True
        where = name + (f" [{label}]" if label else "")
        work = extract_long_entries(df, m["user"], m["date"], m["effort"], m["proj"],
                                    m["task"], m["cust"], warn=warn, where=where)
        if work.empty:
            continue
        sig = pd.util.hash_pandas_object(work, index=False).values.tobytes()
        if sig in sigs:                            # same rows already taken from
            if warn is not None:                   # another tab (pivot / copy)
                warn.append(f"{where}: identical to an earlier sheet, not counted twice.")
            continue
        sigs.add(sig)
        frames.append(work)

    if not mapped_any:
        raise ValueError(
            "expected Workingdate/Effort/User columns; got "
            + " | ".join(f"{lbl or 'sheet'}: {c}" for lbl, c in seen))
    if not frames:
        raise ValueError("matched the expected columns but every row was unusable")

    out = pd.concat(frames, ignore_index=True)
    out["team"] = team_from_filename(name)         # sub team from filename code
    if warn is not None and team_from_filename(name) is None:
        warn.append(f"{name}: no HI0xxxxx cost-centre code in the filename, "
                    f"so these rows have no sub team.")
    return out


def build_data():
    src = OneDriveSource()
    files = src.load_files()
    holidays = load_holidays()

    warns: list[str] = []
    frames = []
    seen_files: dict[str, str] = {}
    for name, year, raw in files:
        # The same filename in the base folder AND a year subfolder used to be
        # loaded twice, doubling that month's hours.
        prev = seen_files.get(name.lower())
        if prev is not None:
            warns.append(f"{name}: also present in {prev or 'the base folder'} — "
                         f"reading the copy in {year or 'the base folder'} only once.")
            continue
        seen_files[name.lower()] = year or ""
        try:
            frames.append(extract_file(name, raw, warn=warns))
        except Exception as e:  # surface the offending file + its real columns
            try:
                cols = list(read_dataframe(name, raw).columns)
            except Exception:
                cols = "<could not read file>"
            raise OneDriveError(
                f"[build {VERSION}] Failed to parse '{name}': {type(e).__name__}: {e}. "
                f"Columns detected: {cols}"
            )

    work = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=["name", "date", "eff", "bill", "task", "proj", "cust"])

    # Group every entry into its 26->25 billing period by Workingdate.
    months = []
    if not work.empty:
        work = work.copy()
        work["period"] = work["date"].map(period_of)
        for (year, month), g in work.groupby("period", sort=True):
            months.append(build_period_month(year, month, g, holidays))
    months.sort(key=lambda mo: (mo["year"], mo["month"]))
    present = {p["team"] for mo in months for p in mo["people"] if p.get("team")}
    subteams = sorted(present, key=lambda t: (SUBTEAM_ORDER.index(t) if t in SUBTEAM_ORDER else 99, t))

    # ---- project-level aggregation (Projects page) ----
    projects = []
    if not work.empty and "proj" in work.columns:
        pj = work["proj"].astype(str).str.strip()
        pw = work[(pj != "") & (pj.str.lower() != "nan")]
        for code, g in pw.groupby("proj", sort=False):
            hours = float(g["eff"].sum())
            if hours <= 0:
                continue
            tasks = {}
            for t, e in zip(g["task"], g["eff"]):
                k = map_task(t)
                tasks[k] = tasks.get(k, 0.0) + float(e)
            ma = tasks.get("Corrective Maintenance", 0) + tasks.get("Preventive Maintenance", 0)
            imp = tasks.get("Implement", 0)
            ptype = "MA" if (ma >= imp and ma > 0) else ("IMP" if imp > 0 else "Other")
            cust = next((c for c in g.get("cust", pd.Series(dtype=str)) if _valid_name(c)), "")
            ppl = []
            for nm, gg in g.groupby("name", sort=True):
                # months: {"YYYY-MM": {task category: hours}} so the client can
                # filter by period AND stack effort by task type
                pmonths = {}
                for d, e, t in zip(gg["date"], gg["eff"], gg["task"]):
                    y, m = period_of(d)
                    key = f"{y}-{m:02d}"
                    cat = map_task(t)
                    cell = pmonths.setdefault(key, {})
                    cell[cat] = round(cell.get(cat, 0.0) + float(e), 1)
                team = None
                if "team" in gg.columns:
                    team = next((t for t in gg["team"] if t and str(t).lower() != "nan"), None)
                ppl.append({"name": nm, "team": team,
                            "hours": round(float(gg["eff"].sum()), 1),
                            "months": pmonths})
            ppl.sort(key=lambda p: -p["hours"])
            projects.append({
                "code": str(code), "customer": str(cust),
                "hours": round(hours, 1),
                "bill": not str(code).upper().startswith(NONBILL_PREFIX),
                "type": ptype,
                "tasks": {k: round(v, 1) for k, v in tasks.items()},
                "people": ppl,
            })
        projects.sort(key=lambda p: -p["hours"])

    for w in warns:
        print(f"[data] {w}", file=sys.stderr, flush=True)

    return {"mode": src.mode, "subteams": subteams, "months": months,
            "projects": projects,
            # anything dropped, zeroed or de-duplicated while reading the
            # workbooks — check this first when totals disagree with Excel
            "warnings": warns,
            "files": sorted(seen_files),
            # full holiday calendar so the client can compute capacity for
            # future months (load simulator forecasts)
            "holidays": sorted(d.isoformat() for d in holidays)}


def get_data(force=False):
    now = time.time()
    if force or _cache["data"] is None or now - _cache["at"] > CACHE_TTL:
        _cache["data"] = build_data()
        _cache["at"] = now
    return _cache["data"]


@app.route("/api/data")
def api_data():
    try:
        return jsonify(get_data(force=request.args.get("refresh") == "1"))
    except OneDriveError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "version": VERSION})


@app.route("/projects")
def projects_page():
    return send_from_directory(app.static_folder, "projects.html")


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8501")))
