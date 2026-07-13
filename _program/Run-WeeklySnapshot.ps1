param(
    [string]$OperationsRoot,
    [string]$InputDir,
    [string]$OutputDir,
    [string]$ArchiveDir
)

$ErrorActionPreference = "Stop"

$ProgramDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = Split-Path -Parent $ProgramDir
if (-not $OperationsRoot) {
    if ((Split-Path -Leaf $RepositoryRoot) -eq "Red Onion Weekly Metrics Automation") {
        $OperationsRoot = Split-Path -Parent $RepositoryRoot
    } else {
        $OperationsRoot = $RepositoryRoot
    }
}
$OperationsRoot = [System.IO.Path]::GetFullPath($OperationsRoot)
$DefaultInputDir = Join-Path $OperationsRoot "01 Daily Reports - Drop Here"
$DefaultOutputDir = Join-Path $OperationsRoot "02 Finished Reports"
$DefaultArchiveDir = Join-Path $OperationsRoot "03 Archive"
$VenvDir = Join-Path $env:LOCALAPPDATA "RedOnionMetrics\.venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

function Get-PythonLauncher {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" *> $null
        if ($LASTEXITCODE -eq 0) {
            return @("py", "-3")
        }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" *> $null
        if ($LASTEXITCODE -eq 0) {
            return @("python")
        }
    }

    throw "Python 3.9 or newer was not found. Install Python from https://www.python.org/downloads/ and check 'Add python.exe to PATH', then rerun this script."
}

if (-not $InputDir) {
    $InputDir = $DefaultInputDir
}
if (-not $OutputDir) {
    $OutputDir = $DefaultOutputDir
}
if (-not $ArchiveDir) {
    $ArchiveDir = $DefaultArchiveDir
}

foreach ($dir in @($InputDir, $OutputDir, $ArchiveDir)) {
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
}

if (-not (Test-Path $VenvPython)) {
    $BasePython = Get-PythonLauncher
    if ($BasePython.Count -gt 1) {
        & $BasePython[0] $BasePython[1..($BasePython.Count - 1)] -m venv $VenvDir
    } else {
        & $BasePython[0] -m venv $VenvDir
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the Red Onion Python environment (exit code $LASTEXITCODE)."
    }
}

& $VenvPython -m pip install --disable-pip-version-check --quiet -r (Join-Path $ProgramDir "requirements.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Could not install the Red Onion program requirements (exit code $LASTEXITCODE)."
}

& $VenvPython (Join-Path $ProgramDir "red_onion_weekly_metrics.py") --input-dir $InputDir --output-dir $OutputDir --archive-dir $ArchiveDir --config (Join-Path $ProgramDir "red_onion_config.json")
$ReportExitCode = $LASTEXITCODE
if ($ReportExitCode -ne 0) {
    exit $ReportExitCode
}

exit 0
