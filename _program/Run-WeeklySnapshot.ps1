param(
    [string]$InputDir,
    [string]$OutputDir,
    [string]$ArchiveDir
)

$ErrorActionPreference = "Stop"

$ProgramDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ProgramDir
$DefaultInputDir = Join-Path $ProjectDir "Daily Reports"
$DefaultOutputDir = Join-Path $ProjectDir "Output"
$DefaultArchiveDir = Join-Path $ProjectDir "Archive - Old Files"
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
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $ProgramDir "requirements.txt")
& $VenvPython (Join-Path $ProgramDir "red_onion_weekly_metrics.py") --input-dir $InputDir --output-dir $OutputDir --archive-dir $ArchiveDir --config (Join-Path $ProgramDir "red_onion_config.json")
