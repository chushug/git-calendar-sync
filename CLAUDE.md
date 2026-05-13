# git-calendar-sync — Dev Context

## Project
Standalone tool: git commits → calendar events. GitHub: https://github.com/chushug/git-calendar-sync

## Architecture
- `gui.py` — Textual TUI entry point (also CLI passthrough: `gui.py sync --method google`)
- `sync.py` — CLI entry point with `generate` / `sync` subcommands
- `core/git_log.py` — git log parsing, `default_ics_path()` (date-stamped, Documents folder)
- `core/models.py` — `Commit` dataclass with stable UID
- `methods/ics.py` — RFC 5545 .ics generation
- `methods/google_cal.py` — Google Calendar API (OAuth2)
- `methods/graph_api.py` — Microsoft Graph API (device-code flow)
- `methods/outlook_com.py` — Classic Outlook COM via PowerShell worker
- `scheduler/_outlook_com_worker.ps1` — PowerShell COM worker (bundled in exe)
- `scheduler/setup_windows.ps1` — Task Scheduler registration (supports exe or python)

## Build
```powershell
.\build.ps1           # builds dist\git-calendar-sync.exe + dist\git-calendar-sync-windows.zip
python -m PyInstaller git-calendar-sync.spec --noconfirm   # manual build
```
Releases: upload `dist\git-calendar-sync-windows.zip` to GitHub Releases manually
(API token doesn't have repo scope; gh CLI needs `gh auth login` first)

## Settings persistence
GUI saves to `~/.git-calendar-sync/settings.json` on every Run.
Credentials (Google file path, Azure client ID) stored there too — no .env needed for exe users.

## Known issues / next steps
- GitHub Release upload is manual (no PAT with repo scope)
- Exe tested on Windows only; macOS/Linux untested
- Classic Outlook COM untested (user uses new Outlook)

## Commit style
```
type(scope): description
```
No `Co-Authored-By` trailers. Ever.
