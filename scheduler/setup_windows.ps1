<#
.SYNOPSIS
    Register a daily Windows Scheduled Task that runs git-calendar-sync.

.PARAMETER Method
    Calendar backend: ics | google | graph | outlook-com

.PARAMETER ProjectDir
    Path to the git-calendar-sync project folder.
    Default: parent directory of this script.

.PARAMETER Time
    Daily trigger time in HH:mm format. Default: 23:30.

.PARAMETER Days
    How many days back to sync each run. Default: 1.

.PARAMETER Uninstall
    Remove the scheduled task.

.EXAMPLE
    .\setup_windows.ps1 -Method google
    .\setup_windows.ps1 -Method graph -Time 22:00
    .\setup_windows.ps1 -Uninstall
#>
param(
    [ValidateSet("ics","google","graph","outlook-com")]
    [string]$Method = "ics",
    [string]$ProjectDir = "",
    [string]$Time = "23:30",
    [int]   $Days = 1,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$TaskFolder = "git-calendar-sync"
$TaskName   = "DailySync"
$FullPath   = "\$TaskFolder\$TaskName"

if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
}
$ProjectDir = (Resolve-Path $ProjectDir).Path

# --- Uninstall ---
if ($Uninstall) {
    if (Get-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskPath "\$TaskFolder\" -TaskName $TaskName -Confirm:$false
        Write-Host "Removed: $FullPath"
    } else {
        Write-Host "Task not found: $FullPath"
    }
    exit 0
}

# --- Find Python 3 ---
function Get-Python3Exe {
    foreach ($name in @("py", "python3", "python")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        $ver = & $cmd.Source --version 2>&1
        if ("$ver" -match "Python 3\.") { return $cmd.Source }
    }
    throw "Python 3 not found on PATH."
}

$pyExe   = Get-Python3Exe
$syncPy  = Join-Path $ProjectDir "sync.py"
$cmdArgs = "/C `"$pyExe`" `"$syncPy`" --method $Method --days $Days"

$principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$action  = New-ScheduledTaskAction -Execute "cmd.exe" -Argument $cmdArgs -WorkingDirectory $ProjectDir

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
    -Description "git-calendar-sync daily sync via --method $Method" | Out-Null

Write-Host ""
Write-Host "Scheduled task registered:"
Write-Host "  Path   : $FullPath"
Write-Host "  Method : $Method"
Write-Host "  Time   : $Time daily"
Write-Host "  Days   : last $Days day(s)"
Write-Host ""
Write-Host "To remove: .\setup_windows.ps1 -Uninstall"
