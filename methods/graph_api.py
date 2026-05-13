"""
Method: graph
Syncs commits to Microsoft Calendar via Microsoft Graph API (OAuth2 device-code flow).
Works with the new Outlook for Windows, Outlook on the web, and Microsoft 365.

Setup (one-time):
    python sync.py --method graph --setup

Requirements:
    pip install msal requests

Azure App Registration (free):
    1. Go to https://portal.azure.com -> Azure Active Directory -> App registrations
    2. New registration: name anything, Accounts in any org + personal (multitenant + personal)
    3. Authentication -> Add platform -> Mobile/desktop -> add https://login.microsoftonline.com/common/oauth2/nativeclient
    4. API permissions -> Microsoft Graph -> Delegated -> Calendars.ReadWrite
    5. Copy the Application (client) ID -> set GRAPH_CLIENT_ID in .env
"""

from __future__ import annotations

import json
import os
from datetime import timedelta
from pathlib import Path

import requests

from core.models import Commit

_TOKEN_DIR  = Path.home() / ".git-calendar-sync"
_TOKEN_FILE = _TOKEN_DIR / "graph_token.json"
_AUTHORITY  = "https://login.microsoftonline.com/common"
_SCOPES     = ["Calendars.ReadWrite", "offline_access"]
_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_DURATION_MIN = 15
_GCS_TAG    = "git-calendar-sync"


def _get_client_id() -> str:
    cid = os.environ.get("GRAPH_CLIENT_ID", "").strip()
    if not cid:
        raise ValueError(
            "GRAPH_CLIENT_ID not set.\n"
            "Add it to your .env file. See methods/graph_api.py for setup instructions."
        )
    return cid


def _load_token() -> dict | None:
    if _TOKEN_FILE.exists():
        return json.loads(_TOKEN_FILE.read_text(encoding="utf-8"))
    return None


def _save_token(token: dict) -> None:
    _TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    _TOKEN_FILE.write_text(json.dumps(token, indent=2), encoding="utf-8")


def _acquire_token(client_id: str) -> str:
    try:
        import msal
    except ImportError:
        raise ImportError("msal not installed. Run: pip install msal")

    app = msal.PublicClientApplication(client_id, authority=_AUTHORITY)

    cached = _load_token()
    if cached:
        result = app.acquire_token_by_refresh_token(cached["refresh_token"], scopes=_SCOPES)
        if "access_token" in result:
            _save_token(result)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow failed: {flow}")

    print("\n" + flow["message"])
    print("Waiting for authentication...\n")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result.get('error_description', result)}")

    _save_token(result)
    print(f"Token saved to {_TOKEN_FILE}")
    return result["access_token"]


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _existing_hashes(token: str) -> set[str]:
    """Return set of git hashes already stored in Graph calendar events."""
    hashes: set[str] = set()
    url = f"{_GRAPH_BASE}/me/events?$select=body&$filter=contains(body/content,'{_GCS_TAG}')&$top=999"
    while url:
        resp = requests.get(url, headers=_headers(token))
        resp.raise_for_status()
        data = resp.json()
        for item in data.get("value", []):
            body = item.get("body", {}).get("content", "")
            for line in body.splitlines():
                if line.startswith("Hash:"):
                    hashes.add(line.split(":", 1)[1].strip())
        url = data.get("@odata.nextLink")
    return hashes


def setup(client_id: str | None = None) -> None:
    cid = client_id or _get_client_id()
    print("Starting Microsoft Graph device-code authentication flow...")
    _acquire_token(cid)
    print("Setup complete.")


def run(commits: list[Commit], repo_name: str, dry_run: bool) -> None:
    if dry_run:
        print(f"[dry-run] Would sync {len(commits)} commits to Microsoft Calendar.")
        for c in commits:
            print(f"  {c}")
        return

    client_id = _get_client_id()
    token = _acquire_token(client_id)
    existing = _existing_hashes(token)
    print(f"Found {len(existing)} already-synced commit(s) in Microsoft Calendar.")

    added = skipped = 0
    for c in commits:
        if c.hash in existing:
            print(f"  Skip (exists): {c.short_hash}  {c.subject}")
            skipped += 1
            continue

        end_dt = c.dt + timedelta(minutes=_DURATION_MIN)
        body_text = (
            f"Hash: {c.hash}\n"
            f"Author: {c.email}\n"
            f"Repo: {repo_name}\n"
            f"Source: {_GCS_TAG}"
        )
        event = {
            "subject": f"[Git] {c.subject}",
            "body": {"contentType": "text", "content": body_text},
            "start": {"dateTime": c.dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "end":   {"dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "showAs": "free",
            "isReminderOn": False,
            "categories": ["Git Commit"],
        }
        resp = requests.post(
            f"{_GRAPH_BASE}/me/events",
            headers=_headers(token),
            json=event,
        )
        resp.raise_for_status()
        print(f"  Added: {c.short_hash}  {c.subject}")
        added += 1

    print(f"Done. Added {added}, skipped {skipped} (already synced).")
