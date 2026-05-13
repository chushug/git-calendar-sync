<#
.SYNOPSIS
    PowerShell worker called by methods/outlook_com.py.
    Reads commits from a JSON temp file and writes them to classic Outlook Calendar via COM.
#>
param(
    [string]$JsonFile,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

if (-not $JsonFile -or -not (Test-Path $JsonFile)) {
    Write-Error "JsonFile not found: $JsonFile"
    exit 1
}

$data     = Get-Content $JsonFile -Raw -Encoding UTF8 | ConvertFrom-Json
$commits  = $data.commits
$repoName = Split-Path $data.repo -Leaf
$category = if ($data.category) { $data.category } else { "Git Commit" }

if ($DryRun) {
    Write-Host "[dry-run] Would create $($commits.Count) Outlook appointment(s):"
    $commits | ForEach-Object { Write-Host "  $($_.hash.Substring(0,7))  $($_.subject)" }
    exit 0
}

# Connect to Outlook
$alreadyRunning = $false
try {
    $outlook        = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $alreadyRunning = $true
    Write-Host "Attached to running Outlook."
} catch {
    $outlook = New-Object -ComObject Outlook.Application
    Write-Host "Started Outlook COM instance."
}

$ns = $outlook.GetNamespace("MAPI")
if (-not $alreadyRunning) {
    try { $ns.Logon("", "", $false, $true) } catch { }
    Start-Sleep -Milliseconds 2000
}

$calendar = $ns.GetDefaultFolder(9)  # olFolderCalendar

# Build dedup set from existing appointments
$existing = @{}
foreach ($item in $calendar.Items) {
    try {
        if ($item.Body -match "Hash:\s*([0-9a-f]{40})") {
            $existing[$Matches[1]] = $true
        }
    } catch { }
}
Write-Host "Existing synced appointments: $($existing.Count)"

$epoch = [datetime]::new(1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
$added = 0

foreach ($c in $commits) {
    if ($existing.ContainsKey($c.hash)) {
        Write-Host "  Skip: $($c.hash.Substring(0,7))  $($c.subject)"
        continue
    }
    $startLocal = $epoch.AddSeconds([long]$c.unix_ts).ToLocalTime()
    $endLocal   = $startLocal.AddMinutes(15)

    $appt             = $outlook.CreateItem(1)
    $appt.Subject     = "[Git] $($c.subject)"
    $appt.Start       = $startLocal
    $appt.End         = $endLocal
    $appt.Body        = "Hash: $($c.hash)`r`nAuthor: $($c.email)`r`nRepo: $repoName"
    $appt.Categories  = $category
    $appt.BusyStatus  = 0      # Free
    $appt.ReminderSet = $false
    $appt.Save()

    Write-Host "  Added: $($c.hash.Substring(0,7))  $($c.subject)"
    $added++
}

Write-Host "Done. Added $added appointment(s)."
exit 0
