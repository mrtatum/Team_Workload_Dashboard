#!/usr/bin/env python3
"""
get_refresh_token.py
====================
Obtain a long-lived refresh token for OneDrive for Business (Microsoft Graph)
using the OAuth 2.0 **device code flow**, then save it to token.json.

Why device code flow?
    - No redirect URI, no local web server, no browser automation needed.
    - You run the script, it prints a short code + URL, you sign in on any
      device, and the script captures the tokens.

Requirements on the Azure AD app registration:
    1. The app must allow public client flows.
       Azure Portal > App registrations > <your app> > Authentication >
       "Advanced settings" > "Allow public client flows" = YES.
    2. Delegated API permissions (Microsoft Graph):
         - offline_access   (required to receive a refresh token)
         - Files.Read.All   (read all files the user can access)
         - User.Read         (basic profile)
       Grant admin consent if your tenant requires it.

Usage:
    pip install -r requirements.txt
    cp .env.example .env        # then fill in CLIENT_ID and TENANT_ID
    python get_refresh_token.py

The refresh token is written to `token.json` (gitignored by default). Copy the
value into your `.env` as REFRESH_TOKEN, or mount token.json and set TOKEN_FILE.
"""

import json
import os
import sys
import time
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # python-dotenv is optional; env vars can be set manually.
    pass

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
CLIENT_ID = os.environ.get("CLIENT_ID", "").strip()
# Use your tenant GUID or domain (e.g. contoso.onmicrosoft.com).
# "organizations" works for any work/school account; "common" allows personal too.
TENANT_ID = os.environ.get("TENANT_ID", "organizations").strip()

# Delegated scopes. offline_access is what gets us the refresh token.
SCOPES = os.environ.get(
    "SCOPES",
    "offline_access Files.Read.All User.Read",
).strip()

TOKEN_FILE = Path(os.environ.get("TOKEN_FILE", "token.json"))

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
DEVICE_CODE_URL = f"{AUTHORITY}/oauth2/v2.0/devicecode"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"


def fail(msg: str) -> None:
    print(f"\nERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def request_device_code() -> dict:
    resp = requests.post(
        DEVICE_CODE_URL,
        data={"client_id": CLIENT_ID, "scope": SCOPES},
        timeout=30,
    )
    if resp.status_code != 200:
        fail(f"device code request failed ({resp.status_code}): {resp.text}")
    return resp.json()


def poll_for_token(device_code: str, interval: int, expires_in: int) -> dict:
    """Poll the token endpoint until the user completes sign-in."""
    deadline = time.time() + expires_in
    while time.time() < deadline:
        resp = requests.post(
            TOKEN_URL,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": CLIENT_ID,
                "device_code": device_code,
            },
            timeout=30,
        )
        data = resp.json()
        if resp.status_code == 200:
            return data

        error = data.get("error")
        if error == "authorization_pending":
            time.sleep(interval)
            continue
        if error == "slow_down":
            interval += 5
            time.sleep(interval)
            continue
        if error == "authorization_declined":
            fail("you declined the sign-in request.")
        if error == "expired_token":
            fail("the device code expired before you signed in. Re-run the script.")
        fail(f"token request failed: {json.dumps(data, indent=2)}")
    fail("timed out waiting for sign-in.")
    return {}  # unreachable


def save_tokens(tokens: dict) -> None:
    payload = {
        "refresh_token": tokens.get("refresh_token"),
        "access_token": tokens.get("access_token"),
        "expires_at": int(time.time()) + int(tokens.get("expires_in", 0)),
        "scope": tokens.get("scope"),
        "client_id": CLIENT_ID,
        "tenant_id": TENANT_ID,
    }
    TOKEN_FILE.write_text(json.dumps(payload, indent=2))
    # Restrict permissions where the OS supports it.
    try:
        os.chmod(TOKEN_FILE, 0o600)
    except OSError:
        pass


def main() -> None:
    if not CLIENT_ID:
        fail("CLIENT_ID is not set. Copy .env.example to .env and fill it in.")

    print("Requesting device code...")
    dc = request_device_code()

    # Show the user the verification instructions.
    print("\n" + "=" * 60)
    print(dc.get("message", "Follow the instructions to sign in."))
    print("=" * 60 + "\n")

    tokens = poll_for_token(
        device_code=dc["device_code"],
        interval=int(dc.get("interval", 5)),
        expires_in=int(dc.get("expires_in", 900)),
    )

    if not tokens.get("refresh_token"):
        fail(
            "No refresh_token returned. Make sure 'offline_access' is in SCOPES "
            "and that the app allows public client flows."
        )

    save_tokens(tokens)
    print(f"\nSuccess! Refresh token saved to: {TOKEN_FILE.resolve()}")
    print("Copy the refresh_token into your .env as REFRESH_TOKEN, or set "
          "TOKEN_FILE to point at this file.")


if __name__ == "__main__":
    main()
