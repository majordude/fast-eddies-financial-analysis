# Reads the redacted statements in <BASE>\Redacted and builds the workbook in <BASE>\Analysis.
#   .\run_analyze.ps1
$ErrorActionPreference = "Stop"
if (-not $env:FEFA_BASE) { $env:FEFA_BASE = "C:\CCardStatements" }
$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { Write-Host "Run setup_windows.ps1 first." -ForegroundColor Red; exit 1 }
& $py (Join-Path $PSScriptRoot "analyze_statements.py")
exit $LASTEXITCODE
