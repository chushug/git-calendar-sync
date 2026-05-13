from __future__ import annotations

import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path

from .models import Commit


def repo_root(start: Path) -> Path:
    result = subprocess.run(
        ["git", "rev-parse", "--show-toplevel"],
        cwd=start, capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Not inside a Git repository: {start}")
    return Path(result.stdout.strip()).resolve()


def _get_branch(repo: Path) -> str:
    result = subprocess.run(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"],
        cwd=repo, capture_output=True, text=True, check=False,
    )
    return result.stdout.strip() if result.returncode == 0 else ""


def _get_remote_url(repo: Path) -> str:
    result = subprocess.run(
        ["git", "remote", "get-url", "origin"],
        cwd=repo, capture_output=True, text=True, check=False,
    )
    return result.stdout.strip() if result.returncode == 0 else ""


def get_commits(repo: Path, since: str, until: str) -> list[Commit]:
    branch = _get_branch(repo)
    remote_url = _get_remote_url(repo)

    cmd = [
        "git", "log",
        "--format=%H\x1f%ae\x1f%at\x1f%s",
        "--no-merges",
        f"--since={since}",
        f"--until={until}",
    ]
    result = subprocess.run(cmd, cwd=repo, capture_output=True, text=True, check=True)
    commits: list[Commit] = []
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split("\x1f", 3)
        if len(parts) < 4:
            continue
        hash_, email, ts_str, subject = parts
        dt = datetime.fromtimestamp(int(ts_str), tz=timezone.utc)
        commits.append(Commit(
            hash=hash_, email=email, dt=dt, subject=subject,
            branch=branch, remote_url=remote_url,
        ))
    return commits


def resolve_date_range(days: int | None, since: str | None, until: str | None) -> tuple[str, str]:
    now = datetime.now(tz=timezone.utc)
    if days is not None:
        s = (now - timedelta(days=days)).strftime("%Y-%m-%d 00:00:00")
        u = now.strftime("%Y-%m-%d %H:%M:%S")
    else:
        s = since or now.strftime("%Y-%m-%d")
        u = until or now.strftime("%Y-%m-%d 23:59:59")
    return s, u


def default_ics_path() -> Path:
    """Platform-appropriate default output path for .ics files."""
    import sys, os
    if sys.platform.startswith("linux"):
        xdg = os.environ.get("XDG_DATA_HOME", "")
        base = Path(xdg) if xdg else Path.home() / ".local" / "share"
        return base / "commit-calendar" / "commits.ics"
    # Windows and macOS: use Documents
    return Path.home() / "Documents" / "CommitCalendar" / "commits.ics"
