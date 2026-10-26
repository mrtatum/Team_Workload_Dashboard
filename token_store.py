"""
token_store.py
==============
Persistent, self-renewing OneDrive refresh token.

Microsoft rotates the refresh token on every use, and a refresh token dies after
90 days WITHOUT use (AADSTS700082). It cannot be made non-expiring. What we can
do is never let it sit unused:

  * Every rotated refresh token is saved to TOKEN_STATE_FILE (a writable
    volume), so a container restart resumes from the NEWEST token instead of
    the original seed in .env.
  * keepalive() renews the token once every TOKEN_KEEPALIVE_HOURS (default 24h)
    from a background thread, even if nobody opens the dashboard. Each renewal
    restarts the 90-day clock, so the chain keeps going indefinitely.
  * Safe across gunicorn workers: an fcntl file lock serialises refreshes and
    writes are atomic (temp file + rename).
  * Re-seeding is automatic: put a new REFRESH_TOKEN in .env and recreate the
    container — the store notices the seed changed and switches to it.

The chain can still break if the account's password changes, IT revokes
sessions, Conditional Access requires a fresh sign-in, or the container is
down for 90+ days. Then: run get_refresh_token.py again and update .env.
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import hashlib
import json
import logging
import os
import tempfile
import threading
import time
from pathlib import Path

import requests

try:
    import fcntl  # POSIX only (Linux container, macOS)
except ImportError:  # pragma: no cover - Windows dev box
    fcntl = None

log = logging.getLogger("token_store")

MAX_INACTIVE_DAYS = 90  # fixed by Microsoft, not configurable


class TokenError(RuntimeError):
    pass


def _fp(token: str) -> str:
    """Short fingerprint so we can detect a changed seed without storing it twice."""
    return hashlib.sha256(token.encode()).hexdigest()[:16] if token else ""


def _iso(ts: float | None) -> str | None:
    if not ts:
        return None
    return _dt.datetime.fromtimestamp(ts, _dt.timezone.utc).isoformat(timespec="seconds")


class TokenStore:
    def __init__(self, client_id: str, tenant_id: str, seed_token: str, scopes: str,
                 state_file: str | os.PathLike | None = None,
                 keepalive_hours: float | None = None) -> None:
        self.client_id = client_id
        self.tenant_id = tenant_id or "organizations"
        self.seed_token = seed_token
        self.seed_fp = _fp(seed_token)
        self.scopes = scopes
        self.token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        self.state_file = Path(state_file or os.environ.get("TOKEN_STATE_FILE", "token_state.json"))
        self.keepalive_hours = float(keepalive_hours if keepalive_hours is not None
                                     else os.environ.get("TOKEN_KEEPALIVE_HOURS", "24"))
        self._thread_lock = threading.Lock()
        self._mem_state: dict = {}           # used when the state file isn't writable
        self.persistent = self._check_writable()
        if not self.persistent:
            log.warning("TOKEN_STATE_FILE %s is not writable — rotated refresh tokens "
                        "will be lost on restart. Mount a writable volume there.",
                        self.state_file)

    # ------------------------------------------------------------ storage
    def _check_writable(self) -> bool:
        try:
            self.state_file.parent.mkdir(parents=True, exist_ok=True)
            return os.access(self.state_file.parent, os.W_OK) and (
                not self.state_file.exists() or os.access(self.state_file, os.W_OK))
        except OSError:
            return False

    @contextlib.contextmanager
    def _file_lock(self):
        if not self.persistent or fcntl is None:
            yield
            return
        lock_path = self.state_file.with_name(self.state_file.name + ".lock")
        with open(lock_path, "a+") as fh:
            fcntl.flock(fh, fcntl.LOCK_EX)
            try:
                yield
            finally:
                fcntl.flock(fh, fcntl.LOCK_UN)

    def _read(self) -> dict:
        if not self.persistent:
            return dict(self._mem_state)
        try:
            return json.loads(self.state_file.read_text())
        except FileNotFoundError:
            return {}
        except (OSError, ValueError) as e:
            log.warning("Could not read %s (%s) — starting from the seed token.",
                        self.state_file, e)
            return {}

    def _write(self, state: dict) -> None:
        if not self.persistent:
            self._mem_state = dict(state)
            return
        fd, tmp = tempfile.mkstemp(dir=str(self.state_file.parent),
                                   prefix=".token_state.", suffix=".tmp")
        try:
            with os.fdopen(fd, "w") as fh:
                json.dump(state, fh, indent=2)
            os.chmod(tmp, 0o600)
            os.replace(tmp, self.state_file)
        except BaseException:
            with contextlib.suppress(OSError):
                os.unlink(tmp)
            raise

    def _usable_state(self, state: dict) -> bool:
        """State belongs to the current seed (a new seed in .env wins)."""
        return bool(state.get("refresh_token")) and state.get("seed_fp") == self.seed_fp \
            and state.get("client_id") == self.client_id

    # ------------------------------------------------------------ refresh
    def _do_refresh(self, state: dict) -> dict:
        rt = state["refresh_token"] if self._usable_state(state) else self.seed_token
        resp = requests.post(self.token_url, data={
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "refresh_token": rt,
            "scope": self.scopes,
        }, timeout=30)
        now = time.time()
        if resp.status_code != 200:
            err = f"Token refresh failed: {resp.status_code} {resp.text}"
            failed = dict(state) if self._usable_state(state) else {
                "seed_fp": self.seed_fp, "client_id": self.client_id,
                "refresh_token": rt, "refreshed_at": None}
            failed.update(last_error=err[:500], last_error_at=now)
            self._write(failed)
            raise TokenError(err)
        data = resp.json()
        new = {
            "client_id": self.client_id,
            "tenant_id": self.tenant_id,
            "seed_fp": self.seed_fp,
            "refresh_token": data.get("refresh_token") or rt,
            "access_token": data["access_token"],
            "access_expires_at": now + int(data.get("expires_in", 3600)),
            "refreshed_at": now,
            "last_error": None,
            "last_error_at": None,
        }
        self._write(new)
        return new

    def access_token(self) -> str:
        """A valid access token; refreshes (and persists the rotation) if needed."""
        with self._thread_lock, self._file_lock():
            state = self._read()
            if (self._usable_state(state) and state.get("access_token")
                    and time.time() < state.get("access_expires_at", 0) - 60):
                return state["access_token"]          # another worker just refreshed
            return self._do_refresh(state)["access_token"]

    def keepalive(self) -> bool:
        """Renew the refresh token if it hasn't been used in keepalive_hours.
        Returns True if a renewal happened."""
        with self._thread_lock, self._file_lock():
            state = self._read()
            last = state.get("refreshed_at") if self._usable_state(state) else None
            if last and time.time() - last < self.keepalive_hours * 3600:
                return False
            self._do_refresh(state)
            log.info("OneDrive refresh token renewed by keep-alive.")
            return True

    def status(self) -> dict:
        state = self._read()
        usable = self._usable_state(state)
        last = state.get("refreshed_at") if usable else None
        return {
            "persistent": self.persistent,
            "state_file": str(self.state_file),
            "last_renewed": _iso(last),
            "expires_if_unused": _iso(last + MAX_INACTIVE_DAYS * 86400) if last else None,
            "keepalive_hours": self.keepalive_hours,
            "last_error": state.get("last_error"),
            "last_error_at": _iso(state.get("last_error_at")),
        }


# ------------------------------------------------------------ env wiring
_store: TokenStore | None = None
_store_lock = threading.Lock()


def store_from_env() -> TokenStore:
    """Build (once per process) a TokenStore from CLIENT_ID/TENANT_ID/REFRESH_TOKEN
    or the TOKEN_FILE written by get_refresh_token.py."""
    global _store
    with _store_lock:
        if _store is not None:
            return _store
        scopes = os.environ.get("SCOPES", "offline_access Files.Read User.Read").strip()
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
                scopes = data.get("scope", scopes)

        # Azure AD does not echo offline_access back in `scope`; without it the
        # refresh returns no new refresh token and the chain dies. Always add it.
        if "offline_access" not in scopes.split():
            scopes = "offline_access " + scopes

        if not client_id or not refresh_token:
            raise TokenError(
                "Missing credentials. Set CLIENT_ID + REFRESH_TOKEN (+TENANT_ID), "
                "mount a token.json at TOKEN_FILE, or set LOCAL_DATA_DIR for testing.")
        _store = TokenStore(client_id, tenant_id, refresh_token, scopes)
        return _store


_keepalive_started = False


def start_keepalive(check_every_seconds: int = 3600) -> None:
    """Background daemon thread: checks hourly, renews once per keepalive_hours.
    Started once per process; the file lock stops gunicorn workers from
    double-renewing."""
    global _keepalive_started
    if _keepalive_started or os.environ.get("LOCAL_DATA_DIR", "").strip():
        return
    try:
        store = store_from_env()
    except TokenError as e:
        log.warning("Token keep-alive not started: %s", e)
        return
    _keepalive_started = True

    def _loop():
        time.sleep(15)  # let the server finish booting
        while True:
            try:
                store.keepalive()
            except Exception as e:  # noqa: BLE001 - never kill the thread
                log.error("Token keep-alive failed: %s", e)
            time.sleep(check_every_seconds)

    threading.Thread(target=_loop, name="token-keepalive", daemon=True).start()
