#!/usr/bin/env python3
"""
git-calendar-sync — Convert git commits into calendar events.

Two top-level commands:

  generate   Write commits to an .ics file (no auth, works anywhere)
  sync       Push commits directly to a calendar service

Examples:
  python sync.py generate --days 7
  python sync.py generate --days 7 --out ~/Desktop/week.ics
  python sync.py sync --method google --days 1
  python sync.py sync --method graph  --setup
  python sync.py sync --method outlook-com --days 1 --dry-run
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv  # type: ignore[import-untyped]
load_dotenv(Path(__file__).parent / ".env")

from core.git_log import get_commits, repo_root, resolve_date_range, default_ics_path


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _add_date_args(p: argparse.ArgumentParser) -> None:
    g = p.add_argument_group("date range")
    g.add_argument("--days", type=int, default=1,
                   help="Sync the last N days (default: 1)")
    g.add_argument("--since", default=None, help="Start date YYYY-MM-DD (ignored if --days is set)")
    g.add_argument("--until", default=None, help="End date/datetime   (ignored if --days is set)")


def _resolve_repo(repo_arg: str | None) -> Path:
    start = Path(repo_arg).resolve() if repo_arg else Path.cwd()
    try:
        return repo_root(start)
    except RuntimeError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


def _fetch_commits(repo: Path, args: argparse.Namespace):
    since, until = resolve_date_range(
        days=args.days,
        since=getattr(args, "since", None),
        until=getattr(args, "until", None),
    )
    print(f"Repo  : {repo}")
    print(f"Range : {since}  ->  {until}")
    try:
        commits = get_commits(repo, since, until)
    except Exception as e:
        print(f"ERROR fetching commits: {e}", file=sys.stderr)
        sys.exit(1)
    if not commits:
        print("No commits found in that range.")
        sys.exit(0)
    print(f"Found {len(commits)} commit(s).")
    return commits


# ---------------------------------------------------------------------------
# Sub-command: generate
# ---------------------------------------------------------------------------

def cmd_generate(args: argparse.Namespace) -> int:
    repo = _resolve_repo(args.repo)
    commits = _fetch_commits(repo, args)

    out_path = Path(args.out) if args.out else default_ics_path()

    from methods.ics import run as ics_run
    ics_run(
        commits,
        repo_name=repo.name,
        out_path=out_path,
        dry_run=args.dry_run,
        duration_minutes=args.duration,
    )
    return 0


# ---------------------------------------------------------------------------
# Sub-command: sync
# ---------------------------------------------------------------------------

def cmd_sync(args: argparse.Namespace) -> int:
    # --setup does not need commits
    if args.setup:
        try:
            if args.method == "google":
                from methods.google_cal import setup as google_setup
                cred_path = Path(
                    args.credentials
                    or os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")
                )
                google_setup(cred_path)
            elif args.method == "graph":
                from methods.graph_api import setup as graph_setup
                graph_setup()
            else:
                print(f"--setup is not used by --method {args.method}.", file=sys.stderr)
                return 1
        except (ImportError, FileNotFoundError, RuntimeError, ValueError) as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return 1
        return 0

    repo = _resolve_repo(args.repo)
    commits = _fetch_commits(repo, args)

    try:
        if args.method == "google":
            from methods.google_cal import run as google_run
            cred_path = Path(
                args.credentials
                or os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")
            )
            google_run(commits, repo.name, cred_path, dry_run=args.dry_run)

        elif args.method == "graph":
            from methods.graph_api import run as graph_run
            graph_run(commits, repo.name, dry_run=args.dry_run)

        elif args.method == "outlook-com":
            from methods.outlook_com import run as com_run
            com_run(commits, repo.name, repo, dry_run=args.dry_run, category=args.category)

    except (ImportError, FileNotFoundError, RuntimeError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    return 0


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(
        prog="sync.py",
        description="Convert git commits into calendar events.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = root.add_subparsers(dest="command", required=True)

    # --- generate ---
    gen = sub.add_parser(
        "generate",
        help="Write commits to an .ics file (no auth required)",
        description="Export commits as an RFC 5545 .ics calendar file.",
    )
    _add_date_args(gen)
    gen.add_argument("--repo", default=None, help="Git repository path (default: cwd)")
    gen.add_argument("--out", default=None,
                     help=f"Output .ics path (default: platform-specific, e.g. {default_ics_path()})")
    gen.add_argument("--duration", type=int, default=15, metavar="MINUTES",
                     help="Event duration in minutes (default: 15)")
    gen.add_argument("--dry-run", action="store_true",
                     help="Print what would be written without creating any file")
    gen.set_defaults(func=cmd_generate)

    # --- sync ---
    syn = sub.add_parser(
        "sync",
        help="Push commits directly to a calendar service",
        description="Sync commits to Google Calendar, Microsoft Calendar, or classic Outlook.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Methods:
  google       Google Calendar API      (requires --setup on first run)
  graph        Microsoft Graph API      (requires --setup on first run; works with new Outlook)
  outlook-com  Classic Outlook via COM  (Windows only; does NOT work with new Outlook)

First-time setup:
  python sync.py sync --method google --setup
  python sync.py sync --method graph  --setup
        """,
    )
    syn.add_argument("--method", required=True,
                     choices=["google", "graph", "outlook-com"],
                     help="Calendar backend")
    _add_date_args(syn)
    syn.add_argument("--repo", default=None, help="Git repository path (default: cwd)")
    syn.add_argument("--credentials", default=None,
                     help="[google] Path to client_secret.json")
    syn.add_argument("--category", default="Git Commit",
                     help="[outlook-com] Outlook category label (default: 'Git Commit')")
    syn.add_argument("--setup", action="store_true",
                     help="Run first-time OAuth flow (google / graph only)")
    syn.add_argument("--dry-run", action="store_true",
                     help="Print what would be created without writing anything")
    syn.set_defaults(func=cmd_sync)

    return root


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
