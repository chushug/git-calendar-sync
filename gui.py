#!/usr/bin/env python3
"""
git-calendar-sync — Textual TUI

Usage:
    python gui.py
    python gui.py --repo /path/to/repo
"""

from __future__ import annotations

import argparse
import io
import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Path setup — works both from source and as a PyInstaller --onefile bundle
# ---------------------------------------------------------------------------

_HERE = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
sys.path.insert(0, str(_HERE))

from dotenv import load_dotenv

# When frozen, look for .env next to the .exe; otherwise next to this file.
_env_path = (Path(sys.executable).parent if getattr(sys, "frozen", False) else _HERE) / ".env"
load_dotenv(_env_path)

from sync import cmd_generate, cmd_sync  # noqa: E402 — must come after path setup


# ---------------------------------------------------------------------------
# Textual app
# ---------------------------------------------------------------------------

from textual import work
from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical
from textual.widgets import (
    Button, Checkbox, Footer, Header, Input, Label,
    RadioButton, RadioSet, RichLog, Rule,
)
from textual.reactive import reactive


METHOD_LABELS = {
    "google":      "Google Calendar          (OAuth2 — first run: choose Setup)",
    "graph":       "Microsoft Graph API      (new Outlook / Microsoft 365)",
    "outlook-com": "Classic Outlook COM      (Windows only, Office 2016+)",
}

ACTION_GENERATE = "generate"
ACTION_SYNC     = "sync"


def _colorize(line: str) -> str:
    if "ERROR" in line or line.lower().startswith("error"):
        return f"[red]{line}[/red]"
    if line.startswith("[dry-run]") or "dry" in line.lower():
        return f"[yellow]{line}[/yellow]"
    if line.startswith("  Added") or "Written" in line:
        return f"[green]{line}[/green]"
    return line


class _LogWriter(io.TextIOBase):
    """Thread-safe stdout/stderr redirector that feeds into a RichLog widget."""

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


class GitCalendarSyncApp(App):
    """TUI for git-calendar-sync."""

    CSS = """
    Screen {
        background: $surface;
    }

    #layout {
        padding: 1 2;
        height: 1fr;
    }

    #left {
        width: 44;
        height: 1fr;
        padding-right: 1;
        border-right: tall $primary-darken-2;
    }

    #right {
        height: 1fr;
        padding-left: 1;
    }

    .section-label {
        color: $text-muted;
        text-style: bold;
        margin-top: 1;
        margin-bottom: 0;
    }

    Input {
        margin-bottom: 1;
    }

    RadioSet {
        border: none;
        margin-bottom: 1;
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
        border: round $primary-darken-2;
        padding: 0 1;
    }

    #log-label {
        color: $text-muted;
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
        self._default_repo = default_repo

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def compose(self) -> ComposeResult:
        yield Header()

        with Horizontal(id="layout"):
            with Vertical(id="left"):
                yield Label("Repository path", classes="section-label")
                yield Input(
                    value=self._default_repo,
                    placeholder="(default: current directory)",
                    id="repo",
                )

                yield Label("Date range", classes="section-label")
                with Horizontal():
                    yield Label("Last  ")
                    yield Input(value="1", id="days", restrict=r"\d*", max_length=3)
                    yield Label("  day(s)")

                yield Rule()

                yield Label("Action", classes="section-label")
                with RadioSet(id="action-set"):
                    yield RadioButton("Generate .ics file", value=True, id="rb-generate")
                    yield RadioButton("Sync to calendar",              id="rb-sync")

                with Container(id="method-box"):
                    yield Label("Calendar method", classes="section-label")
                    with RadioSet(id="method-set"):
                        for mid, mlabel in METHOD_LABELS.items():
                            yield RadioButton(mlabel, id=f"rb-{mid}")

                with Container(id="ics-box"):
                    yield Label("Output path  (leave empty = default)", classes="section-label")
                    yield Input(placeholder="%LOCALAPPDATA%\\CommitCalendar\\commits.ics", id="out")
                    yield Label("Event duration (minutes)", classes="section-label")
                    yield Input(value="15", id="duration", restrict=r"\d*", max_length=3)

                yield Rule()

                yield Checkbox("Dry run  (preview — nothing is written)", id="dry-run")

                yield Button("Run", id="run-btn", variant="primary")

            with Vertical(id="right"):
                yield Label("Output", id="log-label", classes="section-label")
                yield RichLog(id="log", highlight=True, markup=True, wrap=True)

        yield Footer()

    # ------------------------------------------------------------------
    # Reactive / event handlers
    # ------------------------------------------------------------------

    def on_mount(self) -> None:
        self._update_panels()
        self.query_one("#log", RichLog).write("[dim]Ready. Press Run or Ctrl+R to start.[/dim]")

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
        self._do_run()

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
            self.call_from_thread(log.write, f"[dim]Running {ns.command}...[/dim]\n")

            sys.stdout = writer  # type: ignore[assignment]
            sys.stderr = writer  # type: ignore[assignment]

            if ns.command == ACTION_GENERATE:
                rc = cmd_generate(ns)
            else:
                rc = cmd_sync(ns)

            writer.flush()

            if rc == 0:
                self.call_from_thread(log.write, "\n[bold green]Done.[/bold green]")
            else:
                self.call_from_thread(log.write, f"\n[bold red]Exited with code {rc}[/bold red]")

        except SystemExit as e:
            writer.flush()
            code = e.code if isinstance(e.code, int) else 1
            if code == 0:
                self.call_from_thread(log.write, "\n[bold green]Done.[/bold green]")
            else:
                self.call_from_thread(log.write, f"\n[bold red]Exited with code {code}[/bold red]")
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
            repo=repo,
            days=days,
            since=None,
            until=None,
            dry_run=dry_run,
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
            ns.command     = ACTION_SYNC
            ns.method      = method
            ns.credentials = os.environ.get("GOOGLE_CREDENTIALS_FILE", "client_secret.json")
            ns.category    = "Git Commit"
            ns.setup       = False
            ns.func        = cmd_sync

        return ns

    def _selected_method(self) -> str:
        method_set = self.query_one("#method-set", RadioSet)
        idx = method_set.pressed_index
        if idx is None:
            return "google"
        return list(METHOD_LABELS.keys())[idx]


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="git-calendar-sync GUI")
    parser.add_argument("--repo", default="", help="Default repository path")
    args = parser.parse_args()
    GitCalendarSyncApp(default_repo=args.repo).run()


if __name__ == "__main__":
    main()
