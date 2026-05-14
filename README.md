# git-calendar-sync

[![README Chinese](https://img.shields.io/badge/README-Chinese-2563eb?style=for-the-badge)](README.zh.md)
[![README English](https://img.shields.io/badge/README-English-16a34a?style=for-the-badge)](README.md)

Convert git commits into calendar events — automatically, idempotently, and without cluttering your schedule.

## Two ways to use this

### Option 1 — Download the exe (no Python required)

Download `git-calendar-sync-windows.zip` from [Releases](https://github.com/chushug/git-calendar-sync/releases), extract, and double-click **git-calendar-sync.exe**.

The GUI lets you set everything visually - repository path, calendar service, credentials - and remembers your settings between sessions. It now opens in Catppuccin Latte, and the left options panel stays usable in smaller terminal windows with scrolling support. For Google Calendar and Microsoft Graph, enter your credentials directly in the app (no `.env` file needed).

The packaged build does not require Python, but it still needs Git installed so the tool can read commit history.

For **daily auto-sync without opening the GUI**, the exe also works as a CLI:

```
git-calendar-sync.exe sync --method google --days 1 --repo "C:\Projects\MyApp"
```

Register it as a daily Task Scheduler task:

```powershell
.\scheduler\setup_windows.ps1 -Method google -Repo "C:\Projects\MyApp"
```

### Option 2 — Run from source (Python 3.9+)

```bash
pip install -r requirements.txt
python sync.py generate --days 7
python sync.py sync --method google --days 1
python gui.py   # GUI
```

For source installs, credentials can go in a `.env` file. The setup steps below are for this source workflow.

## Features

- Export commits to a standard `.ics` file (no auth, works with any calendar app)
- Sync directly to Google Calendar, Microsoft Calendar (new Outlook), or classic Outlook
- Textual TUI with saved settings, Catppuccin Latte styling, and live output
- Small-window friendly option scrolling with `PageUp` / `PageDown`
- Bundled Windows exe can run either the GUI or headless CLI commands
- Idempotent: running twice never creates duplicates
- Stable UIDs: re-importing the same `.ics` updates events, not doubles them
- Events are always marked **Free** — they never block your calendar
- Daily automation via Windows Task Scheduler or cron

## Compatibility

| Feature | Details |
|---------|---------|
| `.ics` generation | Tested on Windows. Should work on macOS/Linux (untested). |
| Import `.ics` into | Outlook (classic), Apple Calendar, Google Calendar, Thunderbird |
| Direct sync — Google Calendar | Any OS with Python 3.9+ |
| Direct sync — Microsoft Graph | Any OS with Python 3.9+ — works with **new Outlook for Windows**, Outlook on the web, Microsoft 365 |
| Direct sync — Classic Outlook COM | **Windows only**, requires classic Outlook (Office 2016+) |
| New Outlook for Windows (COM) | **Not supported.** Microsoft does not support COM/VBA/VSTO automation in new Outlook. |
| Automated daily sync | Windows Task Scheduler; Linux/macOS via cron |

> **New Outlook note:** The new Outlook for Windows is a web-based app. It does not expose a COM interface.
> Use `--method graph` (Microsoft Graph API) or `--method google` (if your Google account is linked) instead.

## Requirements

### Packaged Windows exe

- Windows
- Git installed and available on `PATH`
- No Python installation required

PowerShell execution policy is only needed if you use the scheduler script or classic Outlook COM:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Running from source

- Python 3.9+
- Git

Install dependencies:

```bash
pip install -r requirements.txt
```

For `.ics` generation only (no Google/Graph sync):

```bash
pip install python-dotenv textual
```

On Windows, PowerShell execution policy is needed for `outlook-com` and the scheduler:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

> Do not use `Unrestricted` — `RemoteSigned` is the minimum needed and is safer.

## Source Quick Start

This section assumes you are running from a source checkout with `python sync.py`.
If you downloaded the exe, use the GUI flow above; the `.env` setup below is only for source installs.

### Option A — Generate `.ics` (no auth)

```bash
python sync.py generate --days 7
# Default output: ~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics  (Windows/macOS)
#                 $XDG_DATA_HOME/commit-calendar/YYYY-MM-DD_commits.ics  (Linux)
#                 ~/.local/share/commit-calendar/YYYY-MM-DD_commits.ics  (Linux fallback)

python sync.py generate --days 7 --out ~/Desktop/commits.ics   # custom path
```

Then import the file into your calendar app by double-clicking it.

> **Important:** Importing a `.ics` file is a one-time snapshot, not a live subscription.
> Re-run the command and re-import to update.
> For continuous sync, use Option B or C.

### Option B — Sync to Google Calendar

**One-time setup:**

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project → enable **Google Calendar API**
3. Create **OAuth 2.0 credentials** (Desktop application) → download the JSON file
4. Copy `.env.example` to `.env` and set:
   ```
   GOOGLE_CREDENTIALS_FILE=path/to/client_secret.json
   ```
5. Run setup (opens browser once):
   ```bash
   python sync.py sync --method google --setup
   ```

**Daily use:**

```bash
python sync.py sync --method google --days 1
```

> **Tip for new Outlook users:** Add your Google account in new Outlook
> (Settings → Accounts → Add account). Google Calendar events appear there automatically.

### Option C — Sync to Microsoft Calendar (new Outlook / Microsoft 365)

Uses the Microsoft Graph API. No COM required. Free to set up.

**One-time Azure setup (free):**

1. Go to [portal.azure.com](https://portal.azure.com) → **App registrations** → New registration
   - Name: `git-calendar-sync` (or anything)
   - Supported account types: **Accounts in any organizational directory + personal Microsoft accounts**
2. Authentication → Add platform → **Mobile and desktop applications**
   - Add: `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. API permissions → Add → Microsoft Graph → Delegated → **Calendars.ReadWrite**
   - Grant admin consent only if your organization requires it
4. Copy the **Application (client) ID** → set in `.env`:
   ```
   GRAPH_CLIENT_ID=your-client-id-here
   ```
5. Run setup (prints a code; visit the URL shown — no browser redirect needed):
   ```bash
   python sync.py sync --method graph --setup
   ```

**Daily use:**

```bash
python sync.py sync --method graph --days 1
```

### Option D — Classic Outlook COM (Windows only)

Requires the classic Microsoft Outlook desktop app (Office 2016 / 2019 / 2021 / Microsoft 365 desktop).
Does **not** work with new Outlook for Windows.

```bash
python sync.py sync --method outlook-com --days 1
python sync.py sync --method outlook-com --days 1 --dry-run   # preview only
```

## All Options

### `generate` — create an .ics file

```
python sync.py generate
  [--days N]          Last N days of commits (default: 1)
  [--since DATE]      Start date YYYY-MM-DD (ignored if --days is set)
  [--until DATE]      End date/datetime     (ignored if --days is set)
  [--repo PATH]       Git repository (default: current directory)
  [--out FILE]        Output .ics path (default: platform-specific)
  [--duration MINS]   Event duration in minutes (default: 15)
  [--dry-run]         Print what would be written; do not create a file
```

### `sync` — push commits to a calendar service

```
python sync.py sync --method google|graph|outlook-com
  [--days N]          Last N days of commits (default: 1)
  [--since DATE]      Start date YYYY-MM-DD
  [--until DATE]      End date/datetime
  [--repo PATH]       Git repository (default: current directory)
  [--credentials FILE] [google] Path to client_secret.json
  [--category LABEL]  [outlook-com] Outlook category (default: "Git Commit")
  [--setup]           Run first-time OAuth flow (google / graph only)
  [--dry-run]         Print what would be created; write nothing
```

## Calendar Event Format

Each commit is mapped to a calendar event as follows:

```
Summary:     [repo-name] commit subject line
Start:       commit timestamp (committer date, UTC)
End:         start + duration (default 15 min)
UID:         <full-hash>@<repo-slug>.commitcal  (stable — never changes)
Description: Hash: <full hash>
             Author: <email>
             Repo: <name>
             Branch: <branch>
             Remote: <origin URL>
Categories:  Git, Commit, <repo-name>
Busy status: Free  (never blocks your calendar)
Reminder:    Off
```

**Duration options:** 5 / 10 / 15 (default) / 30 / 60 minutes via `--duration`.

**UIDs are stable.** The same commit always produces the same UID.
Re-importing the same `.ics` into a compliant calendar app updates the existing event rather than creating a duplicate.

## Automated Daily Sync

### Windows Task Scheduler

```powershell
# Register a daily task at 23:30 using your chosen method
.\scheduler\setup_windows.ps1 -Method google

# Choices: google | graph | outlook-com
.\scheduler\setup_windows.ps1 -Method graph -Time 22:00

# Remove the task
.\scheduler\setup_windows.ps1 -Uninstall
```

> **Note:** This tool is not a background service.
> `setup_windows.ps1` registers a single scheduled task that runs once per day at the specified time.
> Nothing runs in the background between executions.

### macOS / Linux (cron)

```bash
crontab -e
# Add (runs at 23:30 daily):
30 23 * * * cd /path/to/your/repo && python /path/to/git-calendar-sync/sync.py sync --method google --days 1
```

## Deduplication

Running the tool twice for the same date range never creates duplicate events.

| Method | Strategy |
|--------|----------|
| `generate` (`.ics`) | Regenerates the file completely each time. Calendar apps that support UID matching will update existing events on re-import. |
| `google` | Stores `gitHash` in `extendedProperties.private`; queries before inserting. |
| `graph` | Searches for the commit hash in event bodies before inserting. |
| `outlook-com` | Scans existing appointments for `Hash: <sha>` in the body. |

## Output Files

| Platform | Default `.ics` path |
|----------|---------------------|
| Windows  | `~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics` |
| macOS    | `~/Documents/CommitCalendar/YYYY-MM-DD_commits.ics` |
| Linux    | `$XDG_DATA_HOME/commit-calendar/YYYY-MM-DD_commits.ics` (fallback: `~/.local/share/commit-calendar/YYYY-MM-DD_commits.ics`) |

You can always override with `--out /your/path/file.ics`.

> **Tip:** Do not output `.ics` files inside your Git repository directory unless you intend to publish the calendar.

## Security & Privacy

Commit messages, author emails, branch names, and remote URLs are embedded in generated events.

**Before publishing or sharing `.ics` files, verify they do not contain:**
- Private repository names or internal branch names
- Ticket IDs or internal identifiers in commit messages
- Personal email addresses

The GUI stores non-secret preferences in `~/.git-calendar-sync/settings.json`, including the credentials file path and Azure client ID you entered. OAuth token files (`~/.git-calendar-sync/`) and `.env` are listed in `.gitignore` and are never committed.

## Troubleshooting

**`ModuleNotFoundError: No module named 'dotenv'`**
```bash
pip install python-dotenv
```

**`google` method: `FileNotFoundError: client_secret.json`**
Download your OAuth credentials from Google Cloud Console. In the exe GUI, enter the JSON path in the Google credentials field. From source, set `GOOGLE_CREDENTIALS_FILE` in `.env` or pass `--credentials`.

**`graph` method: `GRAPH_CLIENT_ID not set`**
In the exe GUI, enter the Azure Application (client) ID in the Microsoft Graph field. From source, add it to `.env`.

**`outlook-com` method: `Cannot complete the operation. You are not connected.`**
Classic Outlook must be installed and have a configured profile.
New Outlook for Windows is not supported.

**PowerShell script blocked**
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

**Small GUI window hides lower options**
Use the left-panel scrollbar or press `PageUp` / `PageDown`. `Ctrl+Home` and `Ctrl+End` jump to the top and bottom of the options panel.

## Project Structure

```
git-calendar-sync/
├── gui.py                        # Textual TUI entry point
├── sync.py                       # CLI entry point (generate / sync)
├── core/
│   ├── models.py                 # Commit dataclass with stable UID
│   └── git_log.py                # git log parsing, platform paths
├── methods/
│   ├── ics.py                    # generate: RFC 5545 .ics output
│   ├── google_cal.py             # sync: Google Calendar API
│   ├── graph_api.py              # sync: Microsoft Graph API
│   └── outlook_com.py            # sync: Classic Outlook COM (Windows)
├── scheduler/
│   ├── setup_windows.ps1         # Task Scheduler registration
│   └── _outlook_com_worker.ps1   # PowerShell COM worker
├── build.ps1                     # Build Windows .exe release
├── git-calendar-sync.spec        # PyInstaller spec
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md                     # English (this file)
└── README.zh.md                  # Chinese
```

## Known Limitations

- New Outlook for Windows is not supported for COM sync (by design — Microsoft restriction)
- Local `.ics` import is a snapshot, not a live subscription
- Multi-account Outlook is not supported (uses default profile)
- Commits from rebased/amended history are treated as new commits (hash changes)
- Cross-platform `.ics` generation is untested on macOS and Linux

## License

MIT
