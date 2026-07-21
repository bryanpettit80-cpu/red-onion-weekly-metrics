param(
    [string]$OperationsRoot,
    [string]$InputDir,
    [string]$OutputDir,
    [string]$ArchiveDir,
    [switch]$InitializeIntegrityBaseline
)

$ErrorActionPreference = "Stop"

$ProgramDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = Split-Path -Parent $ProgramDir
$IsDeployedCheckout = (Split-Path -Leaf $RepositoryRoot) -eq "Red Onion Weekly Metrics Automation"

function Stop-ReleasePreflight {
    param([Parameter(Mandatory = $true)][string]$Reason)

    throw (
        "Release preflight failed: $Reason " +
        "No reports were created and no source files were moved. " +
        "Ask the technical maintainer to restore a clean, released main checkout."
    )
}

function Invoke-LocalGit {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $PreviousErrorActionPreference = $ErrorActionPreference
    try {
        # Windows PowerShell can promote native stderr to a terminating error when
        # ErrorActionPreference is Stop. Capture the exit code and handle it below.
        $ErrorActionPreference = "Continue"
        $GitOutput = @(& git -C $RepositoryRoot @Arguments 2>$null)
        $GitExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }

    return [pscustomobject]@{
        ExitCode = $GitExitCode
        Text = (($GitOutput | ForEach-Object { $_.ToString() }) -join "`n").Trim()
    }
}

function Assert-DeployedRelease {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Stop-ReleasePreflight "Git is not available, so this deployed checkout cannot be verified."
    }

    $InsideWorkTree = Invoke-LocalGit -Arguments @("rev-parse", "--is-inside-work-tree")
    if ($InsideWorkTree.ExitCode -ne 0 -or $InsideWorkTree.Text -ne "true") {
        Stop-ReleasePreflight "The deployed automation folder is not a Git checkout."
    }

    $StatusResult = Invoke-LocalGit -Arguments @(
        "status", "--porcelain=v1", "--untracked-files=all"
    )
    if ($StatusResult.ExitCode -ne 0) {
        Stop-ReleasePreflight "Git could not inspect the deployed checkout."
    }
    $Status = $StatusResult.Text
    if ($Status) {
        Stop-ReleasePreflight "The deployed automation has local or untracked changes: $Status"
    }

    $BranchResult = Invoke-LocalGit -Arguments @(
        "symbolic-ref", "--quiet", "--short", "HEAD"
    )
    $Branch = $BranchResult.Text
    if ($BranchResult.ExitCode -ne 0 -or $Branch -ne "main") {
        $BranchDescription = if ($Branch) { "'$Branch'" } else { "a detached HEAD" }
        Stop-ReleasePreflight "The deployed automation is on $BranchDescription instead of 'main'."
    }

    $HeadResult = Invoke-LocalGit -Arguments @("rev-parse", "--verify", "HEAD^{commit}")
    if ($HeadResult.ExitCode -ne 0) {
        Stop-ReleasePreflight "Git could not resolve the deployed HEAD commit."
    }
    $Head = $HeadResult.Text

    $OriginMainResult = Invoke-LocalGit -Arguments @(
        "rev-parse", "--verify", "refs/remotes/origin/main^{commit}"
    )
    if ($OriginMainResult.ExitCode -ne 0) {
        Stop-ReleasePreflight "The local origin/main reference is missing."
    }
    $OriginMain = $OriginMainResult.Text
    if ($Head -ne $OriginMain) {
        Stop-ReleasePreflight "HEAD ($Head) does not match local origin/main ($OriginMain)."
    }

    Write-Host "Verified deployed release: main at $($Head.Substring(0, 12))."
}

if ($IsDeployedCheckout) {
    Assert-DeployedRelease
}

if (-not $OperationsRoot) {
    if ($IsDeployedCheckout) {
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

$ProgramArguments = @(
    (Join-Path $ProgramDir "red_onion_weekly_metrics.py"),
    "--input-dir", $InputDir,
    "--output-dir", $OutputDir,
    "--archive-dir", $ArchiveDir,
    "--config", (Join-Path $ProgramDir "red_onion_config.json")
)
if ($InitializeIntegrityBaseline) {
    $ProgramArguments += "--initialize-integrity-baseline"
}

& $VenvPython @ProgramArguments
$ReportExitCode = $LASTEXITCODE
if ($ReportExitCode -ne 0) {
    exit $ReportExitCode
}

exit 0
