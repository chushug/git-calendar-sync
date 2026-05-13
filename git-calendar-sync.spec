# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for git-calendar-sync
# Build:  pyinstaller git-calendar-sync.spec
#    or:  .\build.ps1

from PyInstaller.utils.hooks import collect_all

# Collect all of Textual's CSS, icons, and internal data files
textual_datas, textual_binaries, textual_hiddenimports = collect_all("textual")

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=textual_binaries,
    datas=textual_datas + [
        # Bundle the PowerShell COM worker so outlook-com method works from the exe
        ("scheduler/_outlook_com_worker.ps1", "scheduler"),
        ("scheduler/setup_windows.ps1",       "scheduler"),
    ],
    hiddenimports=textual_hiddenimports + [
        # Dynamically imported inside cmd_generate / cmd_sync
        "methods.ics",
        "methods.google_cal",
        "methods.graph_api",
        "methods.outlook_com",
        # Optional auth libraries — imported lazily; include so the exe works without pip
        "google.auth",
        "google.auth.transport.requests",
        "google_auth_oauthlib.flow",
        "googleapiclient.discovery",
        "msal",
        "requests",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="git-calendar-sync",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,          # Textual TUI needs a terminal — do not set to False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
