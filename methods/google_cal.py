"""
Method: google
Syncs commits to Google Calendar via the Google Calendar API (OAuth2).

Setup (one-time):
    python sync.py --method google --setup

Requirements:
    pip install google-auth-oauthlib google-api-python-client

OAuth credentials:
    1. Go to https://console.cloud.google.com/
    2. Create a project -> Enable Google Calendar API
    3. Create OAuth 2.0 credentials (Desktop app)
    4. Download the JSON and set GOOGLE_CREDENTIALS_FILE in .env
       OR pass --credentials path/to/client_secret.json
"""

from __future__ import annotations

import json
import os
from datetime import timedelta
from pathlib import Path

from core.models import Commit

_TOKEN_DIR = Path.home() / ".git-calendar-sync"
_TOKEN_FILE = _TOKEN_DIR / "google_token.json"
_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
_CALENDAR_ID = "primary"
_DURATION_MIN = 15


def _get_service(credentials_file: Path):
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        raise ImportError(
            "Google API packages not installed.\n"
            "Run: pip install google-auth-oauthlib google-api-python-client"
        )

    creds = None
    if _TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(_TOKEN_FILE), _SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_file.exists():
                raise FileNotFoundError(
                    f"Credentials file not found: {credentials_file}\n"
                    "Download it from Google Cloud Console and set GOOGLE_CREDENTIALS_FILE in .env"
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), _SCOPES)
            creds = flow.run_local_server(port=0)

        _TOKEN_DIR.mkdir(parents=True, exist_ok=True)
        _TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
        print(f"Token saved to {_TOKEN_FILE}")

    return build("calendar", "v3", credentials=creds)


def setup(credentials_file: Path) -> None:
    print("Starting Google Calendar OAuth2 flow...")
    _get_service(credentials_file)
    print("Setup complete. Token saved.")


def _existing_hashes(service) -> set[str]:
    """Fetch git hashes already stored in privateExtendedProperty."""
    hashes: set[str] = set()
    page_token = None
    while True:
        resp = service.events().list(
            calendarId=_CALENDAR_ID,
            privateExtendedProperty="source=git-calendar-sync",
            maxResults=2500,
            pageToken=page_token,
            fields="nextPageToken,items(extendedProperties)",
        ).execute()
        for item in resp.get("items", []):
            props = item.get("extendedProperties", {}).get("private", {})
            h = props.get("gitHash")
            if h:
                hashes.add(h)
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return hashes


def run(commits: list[Commit], repo_name: str, credentials_file: Path, dry_run: bool) -> None:
    if dry_run:
        print(f"[dry-run] Would sync {len(commits)} commits to Google Calendar (primary).")
        for c in commits:
            print(f"  {c}")
        return

    service = _get_service(credentials_file)
    existing = _existing_hashes(service)
    print(f"Found {len(existing)} already-synced commit(s) in Google Calendar.")

    added = skipped = 0
    for c in commits:
        if c.hash in existing:
            print(f"  Skip (exists): {c.short_hash}  {c.subject}")
            skipped += 1
            continue

        end_dt = c.dt + timedelta(minutes=_DURATION_MIN)
        event = {
            "summary": f"[Git] {c.subject}",
            "description": f"Hash: {c.hash}\nAuthor: {c.email}\nRepo: {repo_name}",
            "start": {"dateTime": c.dt.isoformat(), "timeZone": "UTC"},
            "end": {"dateTime": end_dt.isoformat(), "timeZone": "UTC"},
            "transparency": "transparent",
            "colorId": "8",  # graphite
            "extendedProperties": {
                "private": {
                    "source": "git-calendar-sync",
                    "gitHash": c.hash,
                    "repo": repo_name,
                }
            },
        }
        service.events().insert(calendarId=_CALENDAR_ID, body=event).execute()
        print(f"  Added: {c.short_hash}  {c.subject}")
        added += 1

    print(f"Done. Added {added}, skipped {skipped} (already synced).")
