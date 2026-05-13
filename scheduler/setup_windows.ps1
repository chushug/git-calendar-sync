<#
.SYNOPSIS
    Register a daily Windows Scheduled Task that runs git-calendar-sync.

.DESCRIPTION
    Detects whether to use the bundled exe or python sync.py automatically.
    The scheduled task runs silently in the background — no window, no interaction.

.PARAMETER Method
    Calendar backend: ics | google | graph | outlook-com

.PARAMETER Repo
    Path to the git repository to sync. Default: the current directory when
    the task runs (usually the folder this script is in).

.PARAMETER Time
    Daily trigger time in HH:mm format. Default: 23:30.

.PARAMETER Days
    How many days back to sync each run. Default: 1.

.PARAMETER Uninstall
    Remove the scheduled task.

.EXAMPLE
    # Using the downloaded exe (recommended):
    .\scheduler\setup_windows.ps1 -Method google -Repo "C:\Projects\MyApp"

    # Using python (from source):
    .\scheduler\setup_windows.ps1 -Method graph -Time 22:00 -Repo "C:\Projects\MyApp"

    # Generate .ics instead of syncing:
    .\scheduler\setup_windows.ps1 -Method ics -Repo "C:\Projects\MyApp"

    # Remove the task:
    .\scheduler\setup_windows.ps1 -Uninstall
#>
param(
    [ValidateSet("ics", "google", "graph", "outlook-com")]
    [string]$Method = "google",

    [string]$Repo = "",

    [string]$Time = "23:30",

    [int]$Days = 1,

    [switch]$Uninstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$TaskFolder = "git-calendar-sync"
$TaskName   = "DailySync"
$FullPath   = "\$TaskFolder\$TaskName"

# The project root is the parent of this script's folder (scheduler/)
$ProjectDir = Split-Path -Parent $PSScriptRoot

# --- Uninstall ---
if ($Uninstall) {
    if (Get-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -Confirm:$false
        Write-Host "Removed task: $FullPath"
    } else {
        Write-Host "Task not found: $FullPath"
    }
    exit 0
}

# --- Build the sync command ---

# Build subcommand: "ics" maps to "generate", everything else to "sync --method X"
if ($Method -eq "ics") {
    $subArgs = "generate"
} else {
    $subArgs = "sync --method $Method"
}

if ($Days -gt 1) { $subArgs += " --days $Days" }

if ($Repo) {
    $subArgs += " --repo `"$Repo`""
} elseif (-not $Repo) {
    Write-Warning "No -Repo specified. The task will use the current directory at run time."
    Write-Warning "Recommended: .\scheduler\setup_windows.ps1 -Method $Method -Repo `"C:\path\to\your\repo`""
}

# Prefer the bundled exe; fall back to python
$exePath = Join-Path $ProjectDir "git-calendar-sync.exe"

if (Test-Path $exePath) {
    Write-Host "Using exe: $exePath"
    $action = New-ScheduledTaskAction `
        -Execute $exePath `
        -Argument $subArgs `
        -WorkingDirectory $ProjectDir
} else {
    # Find Python 3
    $pyExe = $null
    foreach ($name in @("py", "python3", "python")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        $ver = & $cmd.Source --version 2>&1
        if ("$ver" -match "Python 3\.") { $pyExe = $cmd.Source; break }
    }
    if (-not $pyExe) {
        Write-Error "Python 3 not found and git-calendar-sync.exe is not present in $ProjectDir."
        exit 1
    }

    $syncPy = Join-Path $ProjectDir "sync.py"
    Write-Host "Using Python: $pyExe"
    $action = New-ScheduledTaskAction `
        -Execute $pyExe `
        -Argument "`"$syncPy`" $subArgs" `
        -WorkingDirectory $ProjectDir
}

# --- Register the task ---

$principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

$trigger = New-ScheduledTaskTrigger -Daily -At $Time

$existing = Get-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -Confirm:$false
    Write-Host "Replaced existing task."
}

Register-ScheduledTask `
    -TaskPath    "\$TaskFolder\" `
    -TaskName    $TaskName `
    -Trigger     $trigger `
    -Action      $action `
    -Principal   $principal `
    -Settings    $settings `
    -Description "git-calendar-sync daily $Method sync" | Out-Null

Write-Host ""
Write-Host "Daily sync task registered successfully." -ForegroundColor Green
Write-Host "  Task   : $FullPath"
Write-Host "  Method : $Method"
Write-Host "  Time   : $Time daily"
Write-Host "  Days   : last $Days day(s)"
if ($Repo) { Write-Host "  Repo   : $Repo" }
Write-Host ""
Write-Host "To remove: .\scheduler\setup_windows.ps1 -Uninstall"
