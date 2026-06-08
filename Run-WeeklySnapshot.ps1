param(
    [string]$InputDir,
    [string]$OutputDir
)

$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DefaultSourceDir = Join-Path $ProjectDir "source_daily_reports"
$DefaultOutputDir = Join-Path $ProjectDir "outputs"
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
    $HasDropboxSourceFiles = $false
    if (Test-Path -LiteralPath $DefaultSourceDir) {
        $HasDropboxSourceFiles = [bool](Get-ChildItem -LiteralPath $DefaultSourceDir -Filter "Daily Report*.xls" -File -ErrorAction SilentlyContinue | Select-Object -First 1)
    }

    if ($HasDropboxSourceFiles) {
        $InputDir = $DefaultSourceDir
    } else {
        $InputDir = $ProjectDir
    }
}

if (-not $OutputDir) {
    $OutputDir = $DefaultOutputDir
}

if (-not (Test-Path $VenvPython)) {
    $BasePython = Get-PythonLauncher
    if ($BasePython.Count -gt 1) {
        & $BasePython[0] $BasePython[1..($BasePython.Count - 1)] -m venv $VenvDir
    } else {
        & $BasePython[0] -m venv $VenvDir
    }
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $ProjectDir "requirements.txt")
& $VenvPython (Join-Path $ProjectDir "red_onion_weekly_metrics.py") --input-dir $InputDir --output-dir $OutputDir --config (Join-Path $ProjectDir "red_onion_config.json")
