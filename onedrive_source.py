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
                self.scopes = data.get("scope", self.scopes)

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

    def _list_data(self) -> list[tuple[dict, str | None]]:
        """
        Return [(item, year_folder_name|None), ...]. Files directly in the base
        folder get year=None; files inside a subfolder get that folder's name
        (treated as the year). Matches Excel and CSV.
        """
        folder = self.folder.strip("/")
        base = (f"{GRAPH_BASE}/me/drive/root:/{folder}:/children" if folder
                else f"{GRAPH_BASE}/me/drive/root/children")

        def is_data(it):
            return "file" in it and it["name"].lower().endswith(DATA_EXTS) \
                and not it["name"].startswith("~$")          # skip Excel temp/lock files

        results: list[tuple[dict, str | None]] = []
        for it in self._children(base):
            if is_data(it):
                results.append((it, None))
            elif "folder" in it:
                sub = f"{GRAPH_BASE}/me/drive/items/{it['id']}/children"
                for c in self._children(sub):
                    if is_data(c):
                        results.append((c, it["name"]))
        return results

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
        Return [(filename, year_folder|None, raw_csv_bytes), ...] for every CSV.
        Files may sit directly in the folder, or inside year subfolders.
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
                parent = os.path.dirname(os.path.abspath(p))
                year = None if parent == base else os.path.basename(parent)
                out.append((os.path.basename(p), year, Path(p).read_bytes()))
            if not out:
                raise OneDriveError(f"No Excel/CSV files in LOCAL_DATA_DIR={self.local_dir}")
            return out

        items = self._list_data()
        if not items:
            raise OneDriveError(f"No Excel/CSV files found in '{self.folder}'.")
        return [(it["name"], year, self._download(it)) for it, year in items]

    @property
    def mode(self) -> str:
        return "local" if self.local_dir else "onedrive"
