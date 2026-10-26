# One-time setup for Fast Eddies Financial Spending and Income Analysis (Windows 11).
# Run from this folder:   powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1
# Safe to run again; it skips what is already done.

$ErrorActionPreference = "Stop"
$Here = $PSScriptRoot
$Base = if ($env:FEFA_BASE) { $env:FEFA_BASE } else { "C:\CCardStatements" }

function Say([string]$Text) { Write-Host $Text -ForegroundColor Cyan }
function Fail([string]$Text) { Write-Host $Text -ForegroundColor Red; exit 1 }

Say "1/5  Looking for Python 3.11 to 3.13 ..."
$candidates = @(
    @{ Exe = "py";     Extra = @("-3.12") },
    @{ Exe = "py";     Extra = @("-3.11") },
    @{ Exe = "py";     Extra = @("-3.13") },
    @{ Exe = "python"; Extra = @() }
)
$PyExe = $null
$PyArgs = @()
foreach ($c in $candidates) {
    if (-not (Get-Command $c.Exe -ErrorAction SilentlyContinue)) { continue }
    $extra = $c.Extra
    try {
        $ver = & $c.Exe @extra -c "import sys; print(str(sys.version_info[0]) + '.' + str(sys.version_info[1]))" 2>$null
    } catch { continue }
    if ($LASTEXITCODE -ne 0 -or -not $ver) { continue }
    $parts = ("$ver").Trim().Split(".")
    if ([int]$parts[0] -eq 3 -and [int]$parts[1] -ge 11 -and [int]$parts[1] -le 13) {
        $PyExe = $c.Exe
        $PyArgs = $c.Extra
        Say "     Found Python $ver"
        break
    }
}
if (-not $PyExe) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Say "     Python 3.11 to 3.13 not found. Installing Python 3.12 with winget ..."
        winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
        Fail "Python was installed. Close this PowerShell window, open a new one, and run this setup again."
    }
    Fail "Python 3.11 to 3.13 not found and winget is not available. Install Python 3.12 from python.org/downloads (tick 'Add python.exe to PATH'), then run this setup again."
}

Say "2/5  Creating folders under $Base ..."
foreach ($d in @("Inputs", "Redacted", "Analysis")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Base $d) | Out-Null
}

Say "3/5  Creating the private Python environment (.venv) ..."
$Venv = Join-Path $Here ".venv"
$VenvPy = Join-Path $Venv "Scripts\python.exe"
if (-not (Test-Path $VenvPy)) {
    & $PyExe @PyArgs -m venv $Venv
    if ($LASTEXITCODE -ne 0) { Fail "Could not create the Python environment." }
}

Say "4/5  Installing libraries (this takes a few minutes) ..."
& $VenvPy -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { Fail "pip upgrade failed. Check the internet connection and run this setup again." }
& $VenvPy -m pip install -r (Join-Path $Here "requirements.txt")
if ($LASTEXITCODE -ne 0) { Fail "Library install failed. Check the internet connection and run this setup again." }

Say "5/5  Downloading the spaCy English model (about 400 MB) ..."
& $VenvPy -m spacy download en_core_web_lg
if ($LASTEXITCODE -ne 0) { Fail "Model download failed. Check the internet connection and run this setup again." }

& $VenvPy -c "import pymupdf, openpyxl, presidio_analyzer, spacy; spacy.load('en_core_web_lg'); print('All libraries load correctly.')"
if ($LASTEXITCODE -ne 0) { Fail "A library did not load. Run this setup again; if it repeats, see README_FIRST.txt section 6." }

Write-Host ""
Say "Setup complete."
Write-Host "Put original statement PDFs in $Base\Inputs"
Write-Host "Then, from this folder:"
Write-Host "    .\run_redact.ps1 -Test"
Write-Host "    .\run_redact.ps1"
Write-Host "    .\run_analyze.ps1"
Write-Host "If PowerShell blocks a script, put this in front:  powershell -ExecutionPolicy Bypass -File"
