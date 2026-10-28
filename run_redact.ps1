# Redacts the original statement PDFs in <BASE>\Inputs into <BASE>\Redacted.
#   .\run_redact.ps1 -Test              first PDF only (check the result before running everything)
#   .\run_redact.ps1                    all PDFs directly in Inputs
#   .\run_redact.ps1 BankName           all PDFs in Inputs\BankName (and its subfolders)
#   .\run_redact.ps1 BankName -Test
param(
    [string]$Subfolder = "",
    [switch]$Test
)
$ErrorActionPreference = "Stop"
if (-not $env:FEFA_BASE) { $env:FEFA_BASE = "C:\CCardStatements" }
$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {   # no .venv here: use the python on PATH when it has the packages
    $sys = Get-Command python -ErrorAction SilentlyContinue
    if ($sys) { $py = $sys.Source } else { Write-Host "Run setup_windows.ps1 first." -ForegroundColor Red; exit 1 }
}
$pyArgs = @((Join-Path $PSScriptRoot "presidio_redact_statements.py"))
if ($Subfolder) { $pyArgs += $Subfolder }
if ($Test) { $pyArgs += "--test" }
& $py @pyArgs
exit $LASTEXITCODE
