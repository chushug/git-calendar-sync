#!/usr/bin/env python3
"""
git-calendar-sync — Textual TUI

Usage:
    python gui.py
    python gui.py --repo /path/to/repo
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")


# ---------------------------------------------------------------------------
# Textual app
# ---------------------------------------------------------------------------

from textual import work
from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, ScrollableContainer
from textual.widgets import (
    Button, Checkbox, Footer, Header, Input, Label,
    RadioButton, RadioSet, RichLog, Rule, Static,
)
from textual.reactive import reactive


METHOD_LABELS = {
    "google":      "Google Calendar          (OAuth2 — first run: choose Setup)",
    "graph":       "Microsoft Graph API      (new Outlook / Microsoft 365)",
    "outlook-com": "Classic Outlook COM      (Windows only, Office 2016+)",
}

ACTION_GENERATE = "generate"
ACTION_SYNC     = "sync"


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
            # ---- left panel: settings ----
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

                # method selector — visible only when Sync is selected
                with Container(id="method-box"):
                    yield Label("Calendar method", classes="section-label")
                    with RadioSet(id="method-set"):
                        for mid, mlabel in METHOD_LABELS.items():
                            yield RadioButton(mlabel, id=f"rb-{mid}")

                # extra ics options — visible only when Generate is selected
                with Container(id="ics-box"):
                    yield Label("Output path  (leave empty = default)", classes="section-label")
                    yield Input(placeholder="%LOCALAPPDATA%\\CommitCalendar\\commits.ics", id="out")
                    yield Label("Event duration (minutes)", classes="section-label")
                    yield Input(value="15", id="duration", restrict=r"\d*", max_length=3)

                yield Rule()

                yield Checkbox("Dry run  (preview — nothing is written)", id="dry-run")

                yield Button("Run", id="run-btn", variant="primary")

            # ---- right panel: log output ----
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
    # Background worker
    # ------------------------------------------------------------------

    @work(thread=True)
    def _do_run(self) -> None:
        log  = self.query_one("#log", RichLog)
        btn  = self.query_one("#run-btn", Button)

        try:
            cmd = self._build_command()
            log.write(f"[dim]$ {' '.join(cmd)}[/dim]\n")

            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=str(Path(__file__).parent),
            )
            for line in proc.stdout:  # type: ignore[union-attr]
                line = line.rstrip()
                if "ERROR" in line or "error" in line.lower():
                    self.call_from_thread(log.write, f"[red]{line}[/red]")
                elif line.startswith("[dry-run]") or "dry" in line.lower():
                    self.call_from_thread(log.write, f"[yellow]{line}[/yellow]")
                elif line.startswith("  Added") or "Written" in line:
                    self.call_from_thread(log.write, f"[green]{line}[/green]")
                else:
                    self.call_from_thread(log.write, line)
            proc.wait()

            if proc.returncode == 0:
                self.call_from_thread(log.write, "\n[bold green]Done.[/bold green]")
            else:
                self.call_from_thread(log.write, f"\n[bold red]Exited with code {proc.returncode}[/bold red]")

        except Exception as exc:
            self.call_from_thread(log.write, f"[bold red]Error: {exc}[/bold red]")
        finally:
            self.call_from_thread(setattr, btn, "disabled", False)

    def _build_command(self) -> list[str]:
        sync_py = str(Path(__file__).parent / "sync.py")
        cmd     = [sys.executable, sync_py]

        repo    = self.query_one("#repo",     Input).value.strip()
        days    = self.query_one("#days",     Input).value.strip() or "1"
        dry_run = self.query_one("#dry-run",  Checkbox).value

        if self.action == ACTION_GENERATE:
            cmd.append("generate")
            out      = self.query_one("#out",      Input).value.strip()
            duration = self.query_one("#duration", Input).value.strip() or "15"
            if out:
                cmd += ["--out", out]
            cmd += ["--duration", duration]
        else:
            cmd.append("sync")
            method = self._selected_method()
            cmd += ["--method", method]

        if repo:
            cmd += ["--repo", repo]
        cmd += ["--days", days]
        if dry_run:
            cmd.append("--dry-run")

        return cmd

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
