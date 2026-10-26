"""
onedrive_source.py
==================
Pulls monthly CSV files from a OneDrive for Business folder using a refresh
token. Each CSV is one month (the filename identifies the month).

Credentials (env vars take priority):
  A) Env: CLIENT_ID, TENANT_ID, REFRESH_TOKEN
  B) Mounted token file: TOKEN_FILE=/secrets/token.json (from get_refresh_token.py)
Rotated tokens are persisted and kept alive by token_store.py (TOKEN_STATE_FILE).

Settings:
  ONEDRIVE_FOLDER   folder to read (default "Sysinfra_Workload")
  SCOPES            OAuth scopes (default offline_access Files.Read User.Read)
  LOCAL_DATA_DIR    if set, read *.csv from this folder instead of OneDrive
"""

from __future__ import annotations

import glob
import io
import json
import os
import time
from pathlib import Path

import requests

from token_store import TokenError, store_from_env

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DATA_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")  # Excel preferred; CSV still ok


class OneDriveError(RuntimeError):
    pass


class OneDriveSource:
    def __init__(self) -> None:
        self.local_dir = os.environ.get("LOCAL_DATA_DIR", "").strip()
        self.folder = os.environ.get("ONEDRIVE_FOLDER", "Sysinfra_Workload").strip()
        self.scopes = os.environ.get("SCOPES", "offline_access Files.Read User.Read").strip()
        self._store = None

        if self.local_dir:
            return  # no credentials needed in local mode

        try:
            self._store = store_from_env()
        except TokenError as e:
            raise OneDriveError(str(e)) from None

    # --------------------------------------------------------------- auth
    def _access(self) -> str:
        # TokenStore persists every rotated refresh token (TOKEN_STATE_FILE) so
        # restarts don't fall back to the stale seed in .env — that fallback is
        # what produced AADSTS700082 after 90 days.
        try:
            return self._store.access_token()
        except TokenError as e:
            raise OneDriveError(str(e)) from None

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._access()}"}

    # --------------------------------------------------------------- data
    def _children(self, url: str) -> list[dict]:
        items: list[dict] = []
        while url:
            r = requests.get(url, headers=self._headers(), timeout=30)
            if r.status_code == 404:
                raise OneDriveError(f"Folder '{self.folder}' not found in your OneDrive.")
            r.raise_for_status()
            data = r.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    MAX_DEPTH = 6

    def _base_url(self) -> str:
        folder = self.folder.strip("/")
        return (f"{GRAPH_BASE}/me/drive/root:/{folder}:/children" if folder
                else f"{GRAPH_BASE}/me/drive/root/children")

    @staticmethod
    def _is_data(it) -> bool:
        return ("file" in it and it["name"].lower().endswith(DATA_EXTS)
                and not it["name"].startswith("~$"))   # skip Excel temp/lock files

    def walk(self, include_all: bool = False) -> list[tuple[dict, str, bool]]:
        """
        Walk the whole folder tree. Returns [(item, relative_dir, is_data), ...].

        The previous version descended exactly ONE level, so a layout like
        <folder>/2026/06/timesheet.xlsx was invisible and those months simply
        never appeared on the dashboard. Now recurses to MAX_DEPTH.

        include_all=True also returns files that don't match DATA_EXTS, so a
        diagnostic can show what is being ignored (e.g. an .xlsb export).
        """
        out: list[tuple[dict, str, bool]] = []
        stack = [(self._base_url(), "", 0)]
        while stack:
            url, rel, depth = stack.pop()
            for it in self._children(url):
                if "folder" in it:
                    if depth + 1 <= self.MAX_DEPTH:
                        sub = f"{GRAPH_BASE}/me/drive/items/{it['id']}/children"
                        stack.append((sub, f"{rel}/{it['name']}".strip("/"), depth + 1))
                    continue
                if "file" not in it:
                    continue
                data = self._is_data(it)
                if data or include_all:
                    out.append((it, rel, data))
        out.sort(key=lambda t: (t[1], t[0]["name"]))
        return out

    def _list_data(self) -> list[tuple[dict, str | None]]:
        """
        Return [(item, relative_dir|None), ...] for every Excel/CSV file
        anywhere under the base folder.
        """
        return [(it, rel or None) for it, rel, is_data in self.walk() if is_data]

    def _download(self, item: dict) -> bytes:
        url = item.get("@microsoft.graph.downloadUrl")
        if url:
            r = requests.get(url, timeout=120)
        else:
            r = requests.get(f"{GRAPH_BASE}/me/drive/items/{item['id']}/content",
                             headers=self._headers(), timeout=120)
        r.raise_for_status()
        return r.content

    def load_files(self) -> list[tuple[str, str | None, bytes]]:
        """
        Return [(path, containing_dir|None, raw_bytes), ...] for every Excel/CSV
        file at ANY depth under the folder. `path` is relative to the base
        folder ("2026/06/timesheet.xlsx") so nested files stay distinguishable
        in logs and warnings.
        """
        if self.local_dir:
            base = os.path.abspath(self.local_dir)
            paths = []
            for ext in DATA_EXTS:
                paths += glob.glob(os.path.join(base, "**", "*" + ext), recursive=True)
            out = []
            for p in sorted(set(paths)):
                if os.path.basename(p).startswith("~$"):
                    continue
                rel_dir = os.path.relpath(os.path.dirname(os.path.abspath(p)), base)
                rel_dir = "" if rel_dir == "." else rel_dir.replace(os.sep, "/")
                rel = f"{rel_dir}/{os.path.basename(p)}" if rel_dir else os.path.basename(p)
                out.append((rel, rel_dir or None, Path(p).read_bytes()))
            if not out:
                raise OneDriveError(f"No Excel/CSV files in LOCAL_DATA_DIR={self.local_dir}")
            return out

        items = self._list_data()
        if not items:
            raise OneDriveError(
                f"No Excel/CSV files found anywhere under '{self.folder}'. "
                f"Check the folder name (ONEDRIVE_FOLDER) and that the files are "
                f".xlsx/.xlsm/.xls/.csv — other formats such as .xlsb are ignored. "
                f"Run `python3 diagnose.py --tree` to list what is actually there.")
        return [(f"{rel}/{it['name']}" if rel else it["name"], rel, self._download(it))
                for it, rel in items]

    @property
    def mode(self) -> str:
        return "local" if self.local_dir else "onedrive"
