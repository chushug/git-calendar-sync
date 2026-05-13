<#
.SYNOPSIS
    Build git-calendar-sync.exe using PyInstaller.

.DESCRIPTION
    Installs PyInstaller if missing, runs the build, then zips the result
    into dist\git-calendar-sync-windows.zip for distribution.

.EXAMPLE
    .\build.ps1
    .\build.ps1 -SkipZip     # build exe only, no zip
#>
param(
    [switch]$SkipZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot

function Write-Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# 1. Locate Python 3
# ---------------------------------------------------------------------------
Write-Step "Locating Python 3"

$pyExe = $null
foreach ($candidate in @("python", "python3", "py")) {
    try {
        $ver = & $candidate --version 2>&1
        if ($ver -match "Python 3") {
            $pyExe = $candidate
            Write-Host "Found: $candidate ($ver)"
            break
        }
    } catch { }
}
if (-not $pyExe) {
    Write-Error "Python 3 not found. Install from https://python.org and retry."
    exit 1
}

# ---------------------------------------------------------------------------
# 2. Ensure PyInstaller is installed
# ---------------------------------------------------------------------------
Write-Step "Checking PyInstaller"

$piCheck = & $pyExe -c "import PyInstaller; print(PyInstaller.__version__)" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller not found - installing..."
    & $pyExe -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) { Write-Error "pip install pyinstaller failed"; exit 1 }
} else {
    Write-Host "PyInstaller $piCheck already installed."
}

# ---------------------------------------------------------------------------
# 3. Ensure build dependencies are installed
# ---------------------------------------------------------------------------
Write-Step "Installing project dependencies"
& $pyExe -m pip install -r "$root\requirements.txt"
if ($LASTEXITCODE -ne 0) { Write-Error "pip install -r requirements.txt failed"; exit 1 }

# ---------------------------------------------------------------------------
# 4. Clean previous build artefacts
# ---------------------------------------------------------------------------
Write-Step "Cleaning previous build"
foreach ($dir in @("$root\build", "$root\dist")) {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force
        Write-Host "Removed $dir"
    }
}

# ---------------------------------------------------------------------------
# 5. Run PyInstaller
# ---------------------------------------------------------------------------
Write-Step "Running PyInstaller"
Push-Location $root
try {
    & $pyExe -m PyInstaller "git-calendar-sync.spec" --noconfirm
    if ($LASTEXITCODE -ne 0) { Write-Error "PyInstaller failed"; exit 1 }
} finally {
    Pop-Location
}

$exePath = "$root\dist\git-calendar-sync.exe"
if (-not (Test-Path $exePath)) {
    Write-Error "Build succeeded but exe not found at: $exePath"
    exit 1
}

$sizeMB = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
Write-Host "`nBuilt: $exePath  ($sizeMB MB)" -ForegroundColor Green

if ($SkipZip) { exit 0 }

# ---------------------------------------------------------------------------
# 6. Package into a zip for distribution
# ---------------------------------------------------------------------------
Write-Step "Creating distribution zip"

$zipDir  = "$root\dist\git-calendar-sync-windows"
$zipPath = "$root\dist\git-calendar-sync-windows.zip"

New-Item -ItemType Directory -Path $zipDir -Force | Out-Null

Copy-Item $exePath          "$zipDir\git-calendar-sync.exe"
Copy-Item "$root\README.md" "$zipDir\README.md"
Copy-Item "$root\.env.example" "$zipDir\.env.example" -ErrorAction SilentlyContinue

# Include scheduler scripts for users who want Task Scheduler integration
$schedulerDest = "$zipDir\scheduler"
New-Item -ItemType Directory -Path $schedulerDest -Force | Out-Null
Copy-Item "$root\scheduler\setup_windows.ps1" "$schedulerDest\setup_windows.ps1"

if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path "$zipDir\*" -DestinationPath $zipPath
Remove-Item $zipDir -Recurse -Force

$zipMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "Zipped: $zipPath  ($zipMB MB)" -ForegroundColor Green
Write-Host "`nDone. Share dist\git-calendar-sync-windows.zip with users." -ForegroundColor Green
