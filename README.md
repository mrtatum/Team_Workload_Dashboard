# Team Workload Dashboard

A containerized HTML dashboard that pulls monthly timesheet workbooks from your
OneDrive `Sysinfra_Workload` folder (live, via an OAuth refresh token) and shows
team performance, a per-person daily-hours heatmap, billable vs non-billable
breakdowns, a per-project view, and a job-task-category breakdown — with a
month switcher.

- **Backend:** Flask + gunicorn (`server.py`)
- **Data source:** OneDrive for Business via Microsoft Graph (`onedrive_source.py`), or local files for dev
- **Frontend:** `static/index.html` (Chart.js, dark theme) and `static/projects.html`
- **Port:** `8501`

---

## Contents

- [Quick start (Docker)](#quick-start-docker)
- [Run without Docker (dev)](#run-without-docker-dev)
- [OneDrive / Azure AD setup](#onedrive--azure-ad-setup)
- [Configuration reference (all env vars)](#configuration-reference-all-env-vars)
- [Expected data schema](#expected-data-schema)
- [Job task → category mapping](#job-task--category-mapping)
- [Billable vs non-billable & sub-teams](#billable-vs-non-billable--sub-teams)
- [Periods, capacity & holidays](#periods-capacity--holidays)
- [What the dashboard shows](#what-the-dashboard-shows)
- [HTTP endpoints](#http-endpoints)
- [Deploy anywhere](#deploy-anywhere)
- [Security notes](#security-notes)
- [Project layout](#project-layout)

---

## Quick start (Docker)

```bash
cd team_dashboard
cp .env.example .env       # fill in CLIENT_ID, TENANT_ID, REFRESH_TOKEN
docker compose up --build
# open http://localhost:8501
```

`REFRESH_TOKEN` is obtained once via the OAuth device-code flow — see
[OneDrive / Azure AD setup](#onedrive--azure-ad-setup). Instead of pasting the
token into `.env`, you can mount a `token.json` and set `TOKEN_FILE` — see the
commented lines in `docker-compose.yml`.

## Run without Docker (dev)

```bash
pip install -r requirements.txt

# A) Preview against local files, no OneDrive needed:
LOCAL_DATA_DIR=./sample_data python server.py
# open http://localhost:8501

# B) Against OneDrive (export the 3 credentials first):
export CLIENT_ID=... TENANT_ID=... REFRESH_TOKEN=...
python server.py
```

For production-like serving locally, use the same command as the image:

```bash
gunicorn -b 0.0.0.0:8501 -w 2 --threads 4 -t 120 server:app
```

---

## OneDrive / Azure AD setup

The dashboard reads files from **your own OneDrive for Business** using a
**delegated OAuth refresh token** (device-code flow — no client secret, no
redirect URI, no admin server). You generate the refresh token **once**; the
app then exchanges it for fresh access tokens automatically (Microsoft rotates
the refresh token on each use, so regular usage keeps it alive for up to 90 days
of inactivity).

### 1. Register / configure the Azure AD app

In the [Azure Portal](https://portal.azure.com) → **App registrations**:

1. **Allow public client flows** — *Authentication* → *Advanced settings* →
   "Allow public client flows" = **Yes**. (Device-code is a public-client flow.)
2. **API permissions** (Microsoft Graph, **Delegated**):
   - `offline_access` — **required** to receive a refresh token
   - `Files.Read.All` (or `Files.Read`) — read the files
   - `User.Read` — basic profile
   - Click **Grant admin consent** if your tenant requires it.
3. Note the **Application (client) ID** → this is `CLIENT_ID`.
4. For `TENANT_ID` use your tenant GUID or domain (e.g.
   `contoso.onmicrosoft.com`). `organizations` works for any work/school
   account.

### 2. Generate the refresh token (one time)

Use the included **`get_refresh_token.py`** helper (OAuth 2.0 device-code flow):

```bash
pip install -r requirements.txt
cp .env.example .env          # set CLIENT_ID and TENANT_ID first
python get_refresh_token.py
```

It prints a short code and URL — open the URL, enter the code, and sign in with
the work/school account that **owns** the `Sysinfra_Workload` OneDrive. On
success it writes `token.json` and prints the refresh token. Copy that value
into your `.env` as `REFRESH_TOKEN`, **or** keep `token.json` and point
`TOKEN_FILE` at it.

> If a token is ever revoked or lapses, just re-run the script. If you get
> "No refresh_token returned", confirm `offline_access` is in the scopes and
> that public client flows are enabled on the app.

### 3. Token file alternative

Instead of putting the three values in `.env`, you can supply a JSON file and
point `TOKEN_FILE` at it (handy for mounting as a Kubernetes/Docker secret):

```json
{
  "client_id": "00000000-0000-0000-0000-000000000000",
  "tenant_id": "organizations",
  "refresh_token": "0.AY...",
  "scope": "offline_access Files.Read.All User.Read"
}
```

Env vars take priority; the token file is the fallback when `CLIENT_ID` /
`REFRESH_TOKEN` are not set.

---

## Configuration reference (all env vars)

Copy `.env.example` to `.env`. Everything is configured through environment
variables.

### OneDrive credentials

| Variable        | Default                                  | Purpose                                                        |
|-----------------|------------------------------------------|---------------------------------------------------------------|
| `CLIENT_ID`     | —                                        | Azure app (client) ID. **Required** for OneDrive mode.        |
| `TENANT_ID`     | `organizations`                          | Tenant GUID or domain. `organizations` = any work/school acct.|
| `REFRESH_TOKEN` | —                                        | OAuth refresh token. **Required** for OneDrive mode.          |
| `SCOPES`        | `offline_access Files.Read User.Read`    | OAuth scopes used when refreshing the access token.           |
| `TOKEN_FILE`    | `/secrets/token.json`                    | Fallback JSON credential file if the vars above are absent.   |

### Data source & caching

| Variable                 | Default              | Purpose                                                                 |
|--------------------------|----------------------|-------------------------------------------------------------------------|
| `ONEDRIVE_FOLDER`        | `Sysinfra_Workload`  | OneDrive folder to read (year subfolders supported).                    |
| `LOCAL_DATA_DIR`         | —                    | If set, read `*.xlsx/.xls/.csv` from this folder and **ignore OneDrive** (dev). |
| `CACHE_TTL_SECONDS`      | `300`                | How long parsed data is cached; the **Reload** button bypasses it.      |

### Business rules

| Variable                 | Default   | Purpose                                                                                  |
|--------------------------|-----------|------------------------------------------------------------------------------------------|
| `NONBILL_PROJECT_PREFIX` | `MFE`     | Project Code prefix treated as **non-billable** (case-insensitive). Everything else is billable. |
| `STANDARD_DAY_HOURS`     | `8`       | Normal hours per working day — drives the capacity / normal-hours reference line.         |
| `EXCLUDE_NAMES`          | *(empty)* | Comma-separated user names to drop from all data. Empty = nobody excluded.                |
| `HOLIDAYS`               | *(empty)* | Comma/space-separated ISO dates to exclude from capacity (e.g. `2026-01-01,2026-04-14`). |
| `HOLIDAYS_FILE`          | `holidays.txt` | File of holidays, one ISO date per line (`#` starts a comment). Merged with `HOLIDAYS`.  |

### Runtime

| Variable | Default | Purpose                                              |
|----------|---------|------------------------------------------------------|
| `PORT`   | `8501`  | Port the app listens on (used by `python server.py`).|

> **Sub-teams** are not an env var — they are derived from a `HIxxxxx`
> cost-center code in the **filename**. See
> [Billable vs non-billable & sub-teams](#billable-vs-non-billable--sub-teams).

---

## Expected data schema

Files are **Excel `.xlsx`** (`.xlsm`, `.xls`, `.csv` also accepted). All sheets
are scanned and the one that best matches the timesheet headers is used; a
title/blank row above the table is tolerated. Excel lock files (`~$...`) are
skipped.

The current build reads the **long / transactional** layout — **one row per
work entry**:

| User Name | Workingdate | Effort | Project Code | Job task | Customer |
|-----------|-------------|--------|--------------|----------|----------|

Columns are matched by name (order-tolerant, case/space-insensitive):

- **User Name** — `user name`, `user`, `employee`, `staff`, `resource`, `name`
- **Workingdate** — `workingdate`, `workdate`, or `date` *(required)*
- **Effort** — `effort`, hours for that entry *(required)*
- **Project Code** — `project code`, `project` *(optional; drives billable flag)*
- **Job task** — `job task`, `task type`, `task` *(optional; drives the category breakdown)*
- **Customer** — `customer`, `client` *(optional)*

`Workingdate` and `Effort` and a user column are the minimum required; a file
that lacks them raises a clear parse error.

Files can sit **directly** in the folder or inside **year subfolders**. A single
workbook spanning many months is split automatically into billing periods (see
below), so you don't need one-file-per-month.

```
Sysinfra_Workload/
├── 2025/
│   ├── 05.xlsx
│   └── 06.xlsx
└── 2026/
    └── 01.xlsx
```

> **Note:** an older **wide** layout (one row per person with day columns 26…25)
> was supported in earlier builds but is **not** parsed by the current code,
> which expects the long/transactional schema above.

---

## Job task → category mapping

Each raw **Job task** value from the source data is normalized and mapped to a
**category** used throughout the dashboard. This is the single source of truth —
to add or change a category, edit `TASK_RULES` in `server.py`
(function `map_task`).

**How it works:**

1. The raw value is lower-cased, trimmed, and **all spaces removed** (so spacing
   variants don't matter).
2. Blank / missing / `nan` → **`Others`**.
3. The normalized string is tested against an **ordered** keyword list — the
   **first** keyword that appears as a substring wins. Order matters so specific
   codes match before shorter ones they contain (e.g. `pma`/`cma` before bare
   `ma`).
4. No keyword matches → **`Others`**.

**Current rules (in priority order):**

| # | Keyword (substring, spaces removed) | → Category               | Notes                                            |
|---|-------------------------------------|--------------------------|--------------------------------------------------|
| 1 | `dev,imp,ma,wa`                     | **Presales**             | The "Dev,Imp,MA,WA" presales variant.            |
| 2 | `presale`                          | **Presales**             |                                                  |
| 3 | `meeting`                          | **Meeting**              | Incl. "Meeting/ Admin/ Support-MFE-HI05000*".    |
| 4 | `leave`                            | **Leave**                |                                                  |
| 5 | `standby`                          | **Standby**              |                                                  |
| 6 | `training`                         | **Training**             |                                                  |
| 7 | `pma`                              | **Preventive Maintenance** |                                                |
| 8 | `cma`                              | **Corrective Maintenance** |                                                |
| 9 | `imp`                              | **Implement**            |                                                  |
| 10| `ma` (bare)                        | **Corrective Maintenance** | Fallback for any remaining `*MA*` code.        |
| – | *anything else / blank*            | **Others**               |                                                  |

> **Why order matters:** `pma` and `cma` both contain `ma`. If `ma` came first,
> every preventive/corrective code would be mislabeled. Keep new specific rules
> **above** the shorter ones they contain.

**To extend it**, add a `("keyword", "Category")` tuple to `TASK_RULES` at the
correct priority position and redeploy. Example to route a new "WA" workaround
code to its own category:

```python
TASK_RULES = [
    ("dev,imp,ma,wa", "Presales"),
    ("workaround",    "Workaround"),   # <-- new, placed before bare "ma"
    ("presale",       "Presales"),
    ...
]
```

---

## Billable vs non-billable & sub-teams

- **Billable flag:** a `Project Code` starting with `NONBILL_PROJECT_PREFIX`
  (default `MFE`, case-insensitive) is **non-billable**; everything else is
  **billable**. `% Bill` / `% Non-Bill` are computed from effort, not taken from
  the source.
- **Standby (non-billable):** effort whose category is `Standby` **and** sits on
  a non-billable (`MFE*`) project is also tracked separately as standby
  non-billable hours.
- **Sub-teams:** derived from a `HIxxxxx` cost-center code found in the
  **filename** (e.g. `..._HI05010.xlsx`):

  | Code      | Sub-team   |
  |-----------|------------|
  | `HI05010` | Presales   |
  | `HI05020` | Engineers  |

  Edit `SUBTEAM_MAP` in `server.py` to add codes. Unmapped codes are shown as
  the raw code.

---

## Periods, capacity & holidays

- **Billing periods:** rows are grouped into **26 → 25** periods by
  `Workingdate` (the 26th of the previous month through the 25th of the named
  month). The top-bar **Year** and **Month** dropdowns pick which period to view.
- **Capacity:** working days in the period × `STANDARD_DAY_HOURS` (weekends
  **and** holidays excluded). Weekends are detected from real calendar dates.
- **Normal-hours line:** the "Total hours by person" chart shows a dashed
  reference line at `working days × STANDARD_DAY_HOURS`.
- **Holidays:** list public holidays in `holidays.txt` (one ISO date per line)
  and/or the `HOLIDAYS` env var. Only weekday holidays affect the count. Edit and
  click **Reload** to apply.

---

## What the dashboard shows

- **KPIs:** headcount, total hours, avg billable %, avg non-billable %, peak day load.
- **Heatmap:** people × days, colored by hours (green normal → amber long day → red overload) with hover tooltips, always over the full 26 → 25 calendar.
- **Charts:** total hours by person, billable vs non-billable %, and a job-task-category breakdown for the selected month.
- **Team trend:** total hours + avg billable % across all months.
- **Projects view** at `/projects`.
- **Month switcher** in the top bar; **Reload** forces a fresh OneDrive pull (bypasses the cache).

---

## HTTP endpoints

| Method & path   | Returns                                                            |
|-----------------|-------------------------------------------------------------------|
| `GET /`         | The dashboard (`static/index.html`).                              |
| `GET /projects` | The per-project view (`static/projects.html`).                   |
| `GET /api/data` | Parsed JSON: `months`, `projects`, `subteams`, `holidays`, `mode`.|
| `GET /api/health` | `{"status": "ok", "version": ...}` — used by the healthcheck.   |

---

## Deploy anywhere

A standard single-port (`8501`) container. Push to a registry, then run on
Kubernetes, ECS/Fargate, Cloud Run, Azure Container Apps, Fly, Render, etc.
Supply `CLIENT_ID` / `TENANT_ID` / `REFRESH_TOKEN` (or a mounted `TOKEN_FILE`)
as secrets and expose port `8501`. The image runs gunicorn with 2 workers × 4
threads and a 120s timeout (for OneDrive pulls), and ships a `/api/health`
healthcheck.

---

## Security notes

- **The refresh token is the credential — treat it like a password.** Never
  commit `.env`, `.env.bak`, or `token.json`; they are covered by `.gitignore`.
- This flow uses **no client secret** — only the refresh token.
- Microsoft **rotates** the refresh token on each use; the app keeps the latest
  in memory for the process lifetime.
- The browser needs internet for the Chart.js CDN. To run fully offline, vendor
  `chart.umd.min.js` into `static/` and update the script tag.

---

## Project layout

```
team_dashboard/
├── server.py            # Flask backend: pulls + parses workbooks, serves API + HTML
├── onedrive_source.py   # OneDrive loader (refresh-token auth) + LOCAL_DATA_DIR mode
├── get_refresh_token.py # one-time: device-code sign-in -> token.json
├── build-and-push.sh    # build multi-arch image and push to GHCR
├── static/
│   ├── index.html       # the dashboard (Chart.js, dark theme)
│   └── projects.html    # per-project view
├── sample_data/         # small long-format example for LOCAL_DATA_DIR dev
├── holidays.txt         # public holidays to exclude from capacity
├── Dockerfile           # gunicorn image
├── docker-compose.yml
├── requirements.txt
├── .env.example         # config template (copy to .env)
├── .gitignore
└── .dockerignore
```
