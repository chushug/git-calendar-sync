"""
generate — RFC 5545 .ics file output.  No auth required.

Each commit becomes a VEVENT with:
  - Stable UID  (hash@repo-slug.commitcal) — safe to re-import
  - TRANSP:TRANSPARENT  — never marks you as busy
  - Configurable duration (default 15 min)
"""

from __future__ import annotations

from datetime import timedelta, timezone, datetime
from pathlib import Path

from core.models import Commit


# ---------------------------------------------------------------------------
# RFC 5545 helpers
# ---------------------------------------------------------------------------

def _fmt_dt(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%SZ")


def _escape(text: str) -> str:
    return (
        text.replace("\\", "\\\\")
            .replace(",", "\\,")
            .replace(";", "\\;")
            .replace("\n", "\\n")
    )


def _fold(line: str) -> str:
    """RFC 5545 line folding at 75 octets."""
    encoded = line.encode("utf-8")
    if len(encoded) <= 75:
        return line
    parts: list[bytes] = []
    while len(encoded) > 75:
        cut = 75 if not parts else 74
        parts.append(encoded[:cut])
        encoded = encoded[cut:]
    parts.append(encoded)
    return "\r\n ".join(p.decode("utf-8") for p in parts)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_ics(commits: list[Commit], repo_name: str, duration_minutes: int = 15) -> str:
    now_stamp = _fmt_dt(datetime.now(tz=timezone.utc))
    lines: list[str] = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//git-calendar-sync//{repo_name}//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{repo_name} Git Commits",
    ]
    for c in commits:
        end_dt = c.dt + timedelta(minutes=duration_minutes)

        desc_parts = [
            f"Hash: {c.hash}",
            f"Author: {c.email}",
            f"Repo: {repo_name}",
        ]
        if c.branch:
            desc_parts.append(f"Branch: {c.branch}")
        if c.remote_url:
            desc_parts.append(f"Remote: {c.remote_url}")
        description = _escape("\n".join(desc_parts))

        categories = f"Git,Commit,{repo_name}"

        lines += [
            "BEGIN:VEVENT",
            f"UID:{c.uid}",
            f"DTSTAMP:{now_stamp}",
            f"DTSTART:{_fmt_dt(c.dt)}",
            f"DTEND:{_fmt_dt(end_dt)}",
            f"SUMMARY:[{repo_name}] {_escape(c.subject)}",
            f"DESCRIPTION:{description}",
            f"CATEGORIES:{categories}",
            "TRANSP:TRANSPARENT",
            "STATUS:CONFIRMED",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(l) for l in lines) + "\r\n"


def run(
    commits: list[Commit],
    repo_name: str,
    out_path: Path,
    dry_run: bool,
    duration_minutes: int = 15,
) -> None:
    if dry_run:
        print(f"[dry-run] Would write {len(commits)} events to: {out_path}")
        for c in commits:
            print(f"  {c}")
        return

    ics = build_ics(commits, repo_name, duration_minutes)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(ics, encoding="utf-8")
    print(f"Written: {out_path}  ({len(commits)} events)")
    print()
    print("To import: double-click the .ics file in Outlook, Apple Calendar, or Google Calendar.")
    print("Note: importing creates a one-time snapshot.  Re-run to refresh.")
    print("      For live sync, use: python sync.py sync --method google|graph|outlook-com")
