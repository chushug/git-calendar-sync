#!/usr/bin/env python3
"""
git-calendar-sync — Textual TUI

GUI mode (default):
    python gui.py
    python gui.py --repo /path/to/repo

CLI mode (for Task Scheduler / cron — no GUI window):
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
# CLI passthrough — if first arg is a subcommand, run headless (no GUI)
# This is what Task Scheduler / cron calls.
# ---------------------------------------------------------------------------

def _maybe_run_cli() -> None:
    if len(sys.argv) > 1 and sys.argv[1] in ("generate", "sync"):
        from sync import build_parser
        args = build_parser().parse_args()
        raise SystemExit(args.func(args))

_maybe_run_cli()

# Import after CLI check so Textual doesn't initialise for headless runs
from sync import cmd_generate, cmd_sync  # noqa: E402

from textual import work
from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical
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
    "google":      "Google Calendar          (OAuth2 — requires first-time setup)",
    "graph":       "Microsoft Graph API      (new Outlook / Microsoft 365)",
    "outlook-com": "Classic Outlook COM      (Windows only, Office 2016+)",
}

ACTION_GENERATE = "generate"
ACTION_SYNC     = "sync"


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
    """Captures print() output from sync functions and feeds into RichLog."""

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

    DARK = False  # light theme

    CSS = """
    Screen {
        background: $surface;
    }

    #layout {
        padding: 1 2;
        height: 1fr;
    }

    #left {
        width: 52;
        height: 1fr;
        padding-right: 2;
        border-right: tall $primary;
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

    #method-box {
        display: none;
        margin-bottom: 1;
    }

    #method-box.visible {
        display: block;
    }

    #ics-box {
        display: none;
        margin-bottom: 1;
    }

    #ics-box.visible {
        display: block;
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
        # CLI --repo overrides saved setting
        self._default_repo = default_repo or self._s.get("repo", "")

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def compose(self) -> ComposeResult:  # noqa: PLR0912
        saved_action = self._s.get("action", ACTION_GENERATE)
        saved_method = self._s.get("method", "google")

        yield Header()

        with Horizontal(id="layout"):
            # ── Left panel: settings ──────────────────────────────────
            with Vertical(id="left"):

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

                yield Label("Action", classes="section-label")
                with RadioSet(id="action-set"):
                    yield RadioButton(
                        "Generate .ics file   (drag into any calendar app)",
                        value=(saved_action == ACTION_GENERATE),
                        id="rb-generate",
                    )
                    yield RadioButton(
                        "Sync to calendar     (push commits directly)",
                        value=(saved_action == ACTION_SYNC),
                        id="rb-sync",
                    )

                # Sync options — shown only when "Sync to calendar" is selected
                with Container(id="method-box"):
                    yield Label("Calendar service", classes="section-label")
                    with RadioSet(id="method-set"):
                        for mid, mlabel in METHOD_LABELS.items():
                            yield RadioButton(
                                mlabel,
                                value=(mid == saved_method),
                                id=f"rb-{mid}",
                            )
                    yield Checkbox(
                        "First-time setup  (opens OAuth browser / shows device code)",
                        value=False,
                        id="setup",
                    )

                # Generate options — shown only when "Generate .ics" is selected
                with Container(id="ics-box"):
                    yield Label("Output path", classes="section-label")
                    yield Static(
                        "Leave empty to save to Documents\\CommitCalendar\\commits.ics",
                        classes="help-text",
                    )
                    yield Input(
                        value=self._s.get("out", ""),
                        placeholder="(default: Documents\\CommitCalendar\\commits.ics)",
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
                    "For daily auto-sync: run  scheduler\\setup_windows.ps1 -Method google",
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
        self.action = self._s.get("action", ACTION_GENERATE)
        self._update_panels()
        self.query_one("#log", RichLog).write(
            "[dim]Ready — press [bold]Run[/bold] or [bold]Ctrl+R[/bold] to start.[/dim]\n"
            "[dim]First time? Set your repository path, choose an action, then click Run.[/dim]"
        )

    def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        if event.radio_set.id == "action-set":
            self.action = ACTION_GENERATE if event.index == 0 else ACTION_SYNC
            self._update_panels()

    def _update_panels(self) -> None:
        method_box = self.query_one("#method-box")
        ics_box    = self.query_one("#ics-box")
        if self.action == ACTION_SYNC:
            method_box.add_class("visible")
            ics_box.remove_class("visible")
        else:
            method_box.remove_class("visible")
            ics_box.add_class("visible")

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
            "repo":     self.query_one("#repo",     Input).value,
            "days":     self.query_one("#days",     Input).value,
            "action":   self.action,
            "method":   self._selected_method(),
            "out":      self.query_one("#out",      Input).value,
            "duration": self.query_one("#duration", Input).value,
            "dry_run":  self.query_one("#dry-run",  Checkbox).value,
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
            self.call_from_thread(log.write, f"[dim]Command: {ns.command}[/dim]\n")

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
            ns.command     = ACTION_SYNC
            ns.method      = self._selected_method()
            ns.credentials = os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")
            ns.category    = "Git Commit"
            ns.setup       = self.query_one("#setup", Checkbox).value
            ns.func        = cmd_sync

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
        description="git-calendar-sync GUI — or use 'generate'/'sync' subcommands for CLI mode",
    )
    parser.add_argument("--repo", default="", help="Pre-fill the repository path field")
    args = parser.parse_args()
    GitCalendarSyncApp(default_repo=args.repo).run()


if __name__ == "__main__":
    main()
