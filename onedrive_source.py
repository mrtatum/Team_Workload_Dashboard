"""
onedrive_source.py
==================
Pulls monthly CSV files from a OneDrive for Business folder using a refresh
token. Each CSV is one month (the filename identifies the month).

Credentials (env vars take priority):
  A) Env: CLIENT_ID, TENANT_ID, REFRESH_TOKEN
  B) Mounted token file: TOKEN_FILE=/secrets/token.json (from get_refresh_token.py)

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

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DATA_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")  # Excel preferred; CSV still ok


class OneDriveError(RuntimeError):
    pass


class OneDriveSource:
    def __init__(self) -> None:
        self.local_dir = os.environ.get("LOCAL_DATA_DIR", "").strip()
        self.folder = os.environ.get("ONEDRIVE_FOLDER", "Sysinfra_Workload").strip()
        self.scopes = os.environ.get("SCOPES", "offline_access Files.Read User.Read").strip()
        self._access_token = ""
        self._expires_at = 0.0

        if self.local_dir:
            return  # no credentials needed in local mode

        client_id = os.environ.get("CLIENT_ID", "").strip()
        tenant_id = os.environ.get("TENANT_ID", "organizations").strip()
        refresh_token = os.environ.get("REFRESH_TOKEN", "").strip()

        if not (client_id and refresh_token):
            token_file = Path(os.environ.get("TOKEN_FILE", "/secrets/token.json"))
            if token_file.exists():
                data = json.loads(token_file.read_text())
                client_id = client_id or data.get("client_id", "")
                tenant_id = data.get("tenant_id", tenant_id)
                refresh_token = refresh_token or data.get("refresh_token", "")
                # Azure AD does not echo offline_access back in `scope`, so
                # replaying it verbatim asks for a token with no refresh — the
                # chain then dies at the next rotation. Put it back.
                scope = data.get("scope", self.scopes)
                if "offline_access" not in scope.split():
                    scope = "offline_access " + scope
                self.scopes = scope

        if not client_id or not refresh_token:
            raise OneDriveError(
                "Missing credentials. Set CLIENT_ID + REFRESH_TOKEN (+TENANT_ID), "
                "mount a token.json at TOKEN_FILE, or set LOCAL_DATA_DIR for testing."
            )
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.refresh_token = refresh_token
        self.token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    # --------------------------------------------------------------- auth
    def _access(self) -> str:
        if self._access_token and time.time() < self._expires_at - 60:
            return self._access_token
        resp = requests.post(
            self.token_url,
            data={
                "grant_type": "refresh_token",
                "client_id": self.client_id,
                "refresh_token": self.refresh_token,
                "scope": self.scopes,
            },
            timeout=30,
        )
        if resp.status_code != 200:
            raise OneDriveError(f"Token refresh failed: {resp.status_code} {resp.text}")
        data = resp.json()
        self._access_token = data["access_token"]
        self._expires_at = time.time() + int(data.get("expires_in", 3600))
        if data.get("refresh_token"):
            self.refresh_token = data["refresh_token"]
        return self._access_token

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
