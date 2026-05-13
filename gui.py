#!/usr/bin/env python3
"""
git-calendar-sync — Textual TUI

GUI mode (default):
    python gui.py
    python gui.py --repo /path/to/repo

CLI mode (no GUI — for Task Scheduler / cron):
    python gui.py generate --days 1
    python gui.py sync --method google --days 1 --repo C:\\Projects\\MyApp
"""

from __future__ import annotations

import argparse
import io
import json
import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Path setup — works both from source and as a PyInstaller --onefile bundle
# ---------------------------------------------------------------------------

_HERE = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
sys.path.insert(0, str(_HERE))

from dotenv import load_dotenv

_env_path = (Path(sys.executable).parent if getattr(sys, "frozen", False) else _HERE) / ".env"
load_dotenv(_env_path)

# ---------------------------------------------------------------------------
# CLI passthrough — "generate" / "sync" subcommands run headless (no GUI).
# Task Scheduler calls:  git-calendar-sync.exe sync --method google --days 1
# ---------------------------------------------------------------------------


def _maybe_run_cli() -> None:
    if len(sys.argv) > 1 and sys.argv[1] in ("generate", "sync"):
        from sync import build_parser
        args = build_parser().parse_args()
        raise SystemExit(args.func(args))


_maybe_run_cli()

from sync import cmd_generate, cmd_sync  # noqa: E402

from textual import work
from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, VerticalScroll
from textual.widgets import (
    Button, Checkbox, Footer, Header, Input, Label,
    RadioButton, RadioSet, RichLog, Rule, Static,
)
from textual.reactive import reactive

# ---------------------------------------------------------------------------
# Settings persistence  (~/.git-calendar-sync/settings.json)
# ---------------------------------------------------------------------------

_SETTINGS_PATH = Path.home() / ".git-calendar-sync" / "settings.json"


def _load_settings() -> dict:
    try:
        return json.loads(_SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_settings(s: dict) -> None:
    try:
        _SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
        _SETTINGS_PATH.write_text(json.dumps(s, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

METHOD_LABELS = {
    "google":      "Google Calendar",
    "graph":       "Microsoft Graph API  (new Outlook / Microsoft 365)",
    "outlook-com": "Classic Outlook COM  (Windows only, Office 2016+)",
}

ACTION_GENERATE = "generate"
ACTION_SYNC     = "sync"

# Methods that need no credentials
_NO_CRED_METHODS = {"outlook-com"}


def _colorize(line: str) -> str:
    lo = line.lower()
    if "error" in lo:
        return f"[red]{line}[/red]"
    if line.startswith("[dry-run]") or "dry" in lo:
        return f"[yellow]{line}[/yellow]"
    if "written" in lo or line.startswith("  Added"):
        return f"[green]{line}[/green]"
    if line.startswith("Found ") or line.startswith("Repo") or line.startswith("Range"):
        return f"[cyan]{line}[/cyan]"
    return line


class _LogWriter(io.TextIOBase):
    """Captures print() output from sync functions and streams into RichLog."""

    def __init__(self, app: "GitCalendarSyncApp", log: RichLog) -> None:
        self._app = app
        self._log = log
        self._buf = ""

    def write(self, s: str) -> int:
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._app.call_from_thread(self._log.write, _colorize(line))
        return len(s)

    def flush(self) -> None:
        if self._buf.strip():
            self._app.call_from_thread(self._log.write, _colorize(self._buf))
            self._buf = ""


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

class GitCalendarSyncApp(App):
    """Textual TUI for git-calendar-sync."""

    THEME = "catppuccin-latte"

    CSS = """
    Screen {
        background: $surface;
    }

    #layout {
        padding: 1 2;
        height: 1fr;
    }

    /* Left panel scrolls if content is taller than the terminal */
    VerticalScroll#left {
        width: 54;
        height: 1fr;
        padding-right: 2;
        border-right: tall $primary;
        scrollbar-gutter: stable;
    }

    #right {
        height: 1fr;
        padding-left: 2;
    }

    .section-label {
        color: $primary;
        text-style: bold;
        margin-top: 1;
        margin-bottom: 0;
    }

    .help-text {
        color: $text-muted;
        margin-top: 0;
        margin-bottom: 1;
    }

    .footer-note {
        color: $text-disabled;
        text-style: italic;
        margin-top: 1;
        margin-bottom: 1;
    }

    .days-row {
        height: 3;
        align: left middle;
        margin-bottom: 1;
    }

    .days-row Label {
        margin-top: 1;
    }

    .days-row Input {
        width: 6;
        margin-bottom: 0;
    }

    Input {
        margin-bottom: 1;
    }

    RadioSet {
        border: none;
        margin-bottom: 0;
        padding: 0;
        height: auto;
    }

    #method-box, #ics-box, #cred-google, #cred-graph {
        margin-bottom: 1;
    }

    #run-btn {
        margin-top: 1;
        width: 100%;
    }

    #log {
        height: 1fr;
        border: round $primary;
        padding: 0 1;
    }

    #log-label {
        color: $primary;
        text-style: bold;
        margin-bottom: 0;
    }
    """

    TITLE = "git-calendar-sync"
    SUB_TITLE = "Convert git commits into calendar events"
    BINDINGS = [("ctrl+c", "quit", "Quit"), ("ctrl+r", "run", "Run")]

    action: reactive[str] = reactive(ACTION_GENERATE)

    def __init__(self, default_repo: str = "") -> None:
        super().__init__()
        self._s = _load_settings()
        self._default_repo = default_repo or self._s.get("repo", "")

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def compose(self) -> ComposeResult:
        saved_action = self._s.get("action", ACTION_GENERATE)
        saved_method = self._s.get("method", "google")

        # Pre-fill credentials from saved settings, falling back to env vars
        saved_google_creds = (
            self._s.get("google_creds")
            or os.environ.get("GOOGLE_CREDENTIALS_FILE", "")
        )
        saved_graph_id = (
            self._s.get("graph_client_id")
            or os.environ.get("GRAPH_CLIENT_ID", "")
        )

        yield Header()

        with Horizontal(id="layout"):

            # ── Left panel (scrollable) ───────────────────────────────
            with VerticalScroll(id="left"):

                # Repository path
                yield Label("Repository path", classes="section-label")
                yield Static(
                    "Full path to your git project folder. "
                    "Leave empty to use the current directory.",
                    classes="help-text",
                )
                yield Input(
                    value=self._default_repo,
                    placeholder="e.g.  C:\\Projects\\MyApp",
                    id="repo",
                )

                # Date range
                yield Label("Sync range", classes="section-label")
                with Horizontal(classes="days-row"):
                    yield Label("Last ")
                    yield Input(
                        value=self._s.get("days", "1"),
                        id="days",
                        restrict=r"\d*",
                        max_length=3,
                    )
                    yield Label(" day(s) of commits")

                yield Rule()

                # Action selector
                yield Label("Action", classes="section-label")
                with RadioSet(id="action-set"):
                    yield RadioButton(
                        "Generate .ics file   (import into any calendar app)",
                        value=(saved_action == ACTION_GENERATE),
                        id="rb-generate",
                    )
                    yield RadioButton(
                        "Sync to calendar     (push commits directly)",
                        value=(saved_action == ACTION_SYNC),
                        id="rb-sync",
                    )

                # ── Sync options (hidden until "Sync" is chosen) ──────
                with Container(id="method-box"):
                    yield Label("Calendar service", classes="section-label")
                    with RadioSet(id="method-set"):
                        for mid, mlabel in METHOD_LABELS.items():
                            yield RadioButton(
                                mlabel,
                                value=(mid == saved_method),
                                id=f"rb-{mid}",
                            )

                    # Google credentials (shown only for google method)
                    with Container(id="cred-google"):
                        yield Label("Google credentials file", classes="section-label")
                        yield Static(
                            "Path to client_secret.json downloaded from "
                            "Google Cloud Console → APIs → Credentials.",
                            classes="help-text",
                        )
                        yield Input(
                            value=saved_google_creds,
                            placeholder="C:\\path\\to\\client_secret.json",
                            id="google-creds",
                        )

                    # Graph / Azure credentials (shown only for graph method)
                    with Container(id="cred-graph"):
                        yield Label("Azure Application (client) ID", classes="section-label")
                        yield Static(
                            "From portal.azure.com → App registrations → "
                            "your app → Overview → Application (client) ID.",
                            classes="help-text",
                        )
                        yield Input(
                            value=saved_graph_id,
                            placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                            id="graph-client-id",
                        )

                    yield Checkbox(
                        "First-time setup  (opens OAuth browser / shows device code)",
                        value=False,
                        id="setup",
                    )

                # ── Generate options (hidden until "Generate" is chosen)
                with Container(id="ics-box"):
                    yield Label("Output path", classes="section-label")
                    yield Static(
                        "Where to save the .ics file. "
                        "Leave empty → Documents\\CommitCalendar\\YYYY-MM-DD_commits.ics",
                        classes="help-text",
                    )
                    yield Input(
                        value=self._s.get("out", ""),
                        placeholder="(default: Documents\\CommitCalendar\\YYYY-MM-DD_commits.ics)",
                        id="out",
                    )
                    yield Label("Event duration (minutes)", classes="section-label")
                    yield Input(
                        value=self._s.get("duration", "15"),
                        id="duration",
                        restrict=r"\d*",
                        max_length=3,
                    )

                yield Rule()

                yield Checkbox(
                    "Dry run  (preview — nothing is written or sent)",
                    value=self._s.get("dry_run", False),
                    id="dry-run",
                )
                yield Button("Run", id="run-btn", variant="primary")
                yield Static(
                    "Settings are saved when you click Run.\n"
                    "For daily auto-sync without the GUI:\n"
                    "  scheduler\\setup_windows.ps1 -Method google -Repo \"C:\\your\\repo\"",
                    classes="help-text footer-note",
                )

            # ── Right panel: output log ───────────────────────────────
            with Vertical(id="right"):
                yield Label("Output", id="log-label", classes="section-label")
                yield RichLog(id="log", highlight=True, markup=True, wrap=True)

        yield Footer()

    # ------------------------------------------------------------------
    # Reactive / event handlers
    # ------------------------------------------------------------------

    def on_mount(self) -> None:
        # Hide all conditional panels first, then reveal correct ones
        for wid in ("#method-box", "#ics-box", "#cred-google", "#cred-graph"):
            self.query_one(wid).display = False

        self.action = self._s.get("action", ACTION_GENERATE)
        self._update_panels()
        self.query_one("#log", RichLog).write(
            "[dim]Ready — press [bold]Run[/bold] or [bold]Ctrl+R[/bold] to start.[/dim]\n"
            "[dim]First time? Set your repository path, pick an action, then click Run.[/dim]"
        )

    def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        if event.radio_set.id == "action-set":
            self.action = ACTION_GENERATE if event.index == 0 else ACTION_SYNC
            self._update_panels()
        elif event.radio_set.id == "method-set":
            self._update_cred_panel()

    def _update_panels(self) -> None:
        is_sync = self.action == ACTION_SYNC
        self.query_one("#method-box").display = is_sync
        self.query_one("#ics-box").display    = not is_sync
        if is_sync:
            self._update_cred_panel()

    def _update_cred_panel(self) -> None:
        method = self._selected_method()
        self.query_one("#cred-google").display = (method == "google")
        self.query_one("#cred-graph").display  = (method == "graph")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "run-btn":
            self.action_run()

    def action_run(self) -> None:
        log = self.query_one("#log", RichLog)
        log.clear()
        log.write("[bold cyan]Starting...[/bold cyan]")
        self.query_one("#run-btn", Button).disabled = True
        self._persist_settings()
        self._do_run()

    def _persist_settings(self) -> None:
        _save_settings({
            "repo":            self.query_one("#repo",           Input).value,
            "days":            self.query_one("#days",           Input).value,
            "action":          self.action,
            "method":          self._selected_method(),
            "google_creds":    self.query_one("#google-creds",   Input).value,
            "graph_client_id": self.query_one("#graph-client-id", Input).value,
            "out":             self.query_one("#out",            Input).value,
            "duration":        self.query_one("#duration",       Input).value,
            "dry_run":         self.query_one("#dry-run",        Checkbox).value,
        })

    # ------------------------------------------------------------------
    # Background worker — calls sync functions directly (PyInstaller safe)
    # ------------------------------------------------------------------

    @work(thread=True)
    def _do_run(self) -> None:
        log = self.query_one("#log", RichLog)
        btn = self.query_one("#run-btn", Button)

        writer     = _LogWriter(self, log)
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        try:
            ns = self._build_namespace()
            self.call_from_thread(log.write, f"[dim]Running: {ns.command}[/dim]\n")

            sys.stdout = writer  # type: ignore[assignment]
            sys.stderr = writer  # type: ignore[assignment]

            rc = cmd_generate(ns) if ns.command == ACTION_GENERATE else cmd_sync(ns)
            writer.flush()

            if rc == 0:
                self.call_from_thread(log.write, "\n[bold green]Done.[/bold green]")
            else:
                self.call_from_thread(log.write, f"\n[bold red]Failed (exit {rc})[/bold red]")

        except SystemExit as e:
            writer.flush()
            code = e.code if isinstance(e.code, int) else 1
            msg  = "[bold green]Done.[/bold green]" if code == 0 else f"[bold red]Exited with code {code}[/bold red]"
            self.call_from_thread(log.write, f"\n{msg}")
        except Exception as exc:
            writer.flush()
            self.call_from_thread(log.write, f"[bold red]Error: {exc}[/bold red]")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.call_from_thread(setattr, btn, "disabled", False)

    def _build_namespace(self) -> argparse.Namespace:
        repo    = self.query_one("#repo",    Input).value.strip() or None
        days    = int(self.query_one("#days", Input).value.strip() or "1")
        dry_run = self.query_one("#dry-run", Checkbox).value

        ns = argparse.Namespace(
            repo=repo, days=days, since=None, until=None, dry_run=dry_run,
        )

        if self.action == ACTION_GENERATE:
            out      = self.query_one("#out",      Input).value.strip() or None
            duration = int(self.query_one("#duration", Input).value.strip() or "15")
            ns.command  = ACTION_GENERATE
            ns.out      = out
            ns.duration = duration
            ns.func     = cmd_generate
        else:
            method = self._selected_method()
            ns.command  = ACTION_SYNC
            ns.method   = method
            ns.category = "Git Commit"
            ns.setup    = self.query_one("#setup", Checkbox).value
            ns.func     = cmd_sync

            if method == "google":
                creds = self.query_one("#google-creds", Input).value.strip()
                ns.credentials = creds or os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")
            else:
                ns.credentials = os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")

            if method == "graph":
                client_id = self.query_one("#graph-client-id", Input).value.strip()
                if client_id:
                    os.environ["GRAPH_CLIENT_ID"] = client_id

        return ns

    def _selected_method(self) -> str:
        try:
            idx = self.query_one("#method-set", RadioSet).pressed_index
            if idx is not None:
                return list(METHOD_LABELS.keys())[idx]
        except Exception:
            pass
        return self._s.get("method", "google")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "git-calendar-sync GUI. "
            "Pass 'generate' or 'sync' as the first argument for headless CLI mode."
        ),
    )
    parser.add_argument("--repo", default="", help="Pre-fill the repository path field")
    args = parser.parse_args()
    GitCalendarSyncApp(default_repo=args.repo).run()


if __name__ == "__main__":
    main()
