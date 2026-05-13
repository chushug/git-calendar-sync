"""
Method: outlook-com
Syncs commits to classic Outlook (desktop/Office) via Windows COM automation.
Delegates to a PowerShell script; Windows only.

Requirements:
  - Windows OS
  - Classic Microsoft Outlook (Office 2016 / 2019 / 2021 / Microsoft 365 desktop)
  - The new Outlook for Windows does NOT support COM automation
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from core.models import Commit

_PS1 = Path(__file__).parent.parent / "scheduler" / "_outlook_com_worker.ps1"


def _check_platform() -> None:
    if sys.platform != "win32":
        raise RuntimeError("--method outlook-com is only supported on Windows.")


def run(commits: list[Commit], repo_name: str, repo_path: Path, dry_run: bool, category: str = "Git Commit") -> None:
    _check_platform()

    if not _PS1.exists():
        raise FileNotFoundError(f"PowerShell worker not found: {_PS1}")

    payload = [
        {
            "hash": c.hash,
            "email": c.email,
            "unix_ts": int(c.dt.timestamp()),
            "subject": c.subject,
        }
        for c in commits
    ]

    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump({"commits": payload, "repo": str(repo_path), "category": category}, f)
        tmp_path = f.name

    try:
        args = [
            "powershell", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", str(_PS1),
            "-JsonFile", tmp_path,
        ]
        if dry_run:
            args.append("-DryRun")

        result = subprocess.run(args, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"PowerShell worker exited with code {result.returncode}")
    finally:
        Path(tmp_path).unlink(missing_ok=True)
