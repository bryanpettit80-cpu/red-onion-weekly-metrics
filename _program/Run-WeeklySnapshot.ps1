param(
    [string]$OperationsRoot,
    [string]$InputDir,
    [string]$OutputDir,
    [string]$ArchiveDir,
    [switch]$InitializeIntegrityBaseline,
    [switch]$RebuildEnvironment,
    [switch]$HealthCheck,
    [switch]$NonInteractive,
    [switch]$ReviewCorrections,
    [string]$CorrectionsFile,
    [string]$ManagementWorkbook,
    [string]$RebindRestoredIntegrityAnchor
)

$ErrorActionPreference = "Stop"
if ($NonInteractive -and $ReviewCorrections) {
    throw "-ReviewCorrections needs an interactive review; omit -NonInteractive or supply an already reviewed -CorrectionsFile."
}
if ($HealthCheck -and $RebindRestoredIntegrityAnchor) {
    throw "-HealthCheck and -RebindRestoredIntegrityAnchor are separate maintenance operations."
}

$ProgramDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = Split-Path -Parent $ProgramDir
# The runner uses the installed source and locked dependencies. Git branch,
# local notes, and Python cache files are not reporting prerequisites.
$env:RED_ONION_VERIFIED_RELEASE_COMMIT = $null

# Every Python process uses isolated mode so the caller's working directory and
# Python environment variables cannot inject imports. -B prevents this run from
# creating source-tree bytecode. The source-program bootstrap below additionally
# rejects sourceless bytecode imports after the reviewed program path is added.
$env:PYTHONDONTWRITEBYTECODE = "1"
$env:PYTHONPYCACHEPREFIX = Join-Path (
    [System.IO.Path]::GetTempPath()
) ("RedOnionMetrics-PythonCache-" + [guid]::NewGuid().ToString("N"))

function Test-OperatorFolders {
    param([Parameter(Mandatory = $true)][string]$Root)

    foreach ($Name in @("01 Daily Reports - Drop Here", "02 Finished Reports", "03 Archive")) {
        if (Test-Path -LiteralPath (Join-Path $Root $Name) -PathType Container) {
            return $true
        }
    }
    return $false
}

if (-not $OperationsRoot) {
    $ParentRoot = Split-Path -Parent $RepositoryRoot
    if (Test-OperatorFolders -Root $RepositoryRoot) {
        $OperationsRoot = $RepositoryRoot
    } elseif ($ParentRoot -and (Test-OperatorFolders -Root $ParentRoot)) {
        $OperationsRoot = $ParentRoot
    } else {
        $OperationsRoot = $RepositoryRoot
    }
}
$OperationsRoot = [System.IO.Path]::GetFullPath($OperationsRoot)
$DefaultInputDir = Join-Path $OperationsRoot "01 Daily Reports - Drop Here"
$DefaultOutputDir = Join-Path $OperationsRoot "02 Finished Reports"
$DefaultArchiveDir = Join-Path $OperationsRoot "03 Archive"
$StateDir = Join-Path $env:LOCALAPPDATA "RedOnionMetrics"
$VenvDir = Join-Path $StateDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$EnvironmentStatePath = Join-Path $StateDir "environment-state.json"
$IntegrityAnchorDir = Join-Path $env:LOCALAPPDATA "RedOnionMetrics\integrity-anchors"
$ConfigPath = Join-Path $ProgramDir "red_onion_config.json"
$ConfigValidatorPath = Join-Path $ProgramDir "red_onion_config.py"
$ProgramPath = Join-Path $ProgramDir "red_onion_weekly_metrics.py"
$DirectRequirementsPath = Join-Path $ProgramDir "requirements.txt"
$LockedRequirementsPath = Join-Path $ProgramDir "requirements.lock"
$InstallRequirementsPath = if (Test-Path -LiteralPath $LockedRequirementsPath) {
    $LockedRequirementsPath
} else {
    $DirectRequirementsPath
}

function Get-PythonLauncher {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        foreach ($PythonSelector in @("-3.12", "-3.11", "-3.10")) {
            & py $PythonSelector -I -B -c "import sys; raise SystemExit(0 if (3, 10) <= sys.version_info[:2] < (3, 13) else 1)" *> $null
            if ($LASTEXITCODE -eq 0) {
                return @("py", $PythonSelector, "-I", "-B")
            }
        }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -I -B -c "import sys; raise SystemExit(0 if (3, 10) <= sys.version_info[:2] < (3, 13) else 1)" *> $null
        if ($LASTEXITCODE -eq 0) {
            return @("python", "-I", "-B")
        }
    }

    throw "Python 3.10-3.12 was not found. Install a supported Python release from https://www.python.org/downloads/ and check 'Add python.exe to PATH', then rerun this script."
}

function Invoke-PythonLauncher {
    param(
        [Parameter(Mandatory = $true)][string[]]$Launcher,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $Executable = $Launcher[0]
    $PrefixArguments = @($Launcher | Select-Object -Skip 1)
    $CommandOutput = @(& $Executable @PrefixArguments @Arguments)
    $ExitCode = $LASTEXITCODE
    foreach ($OutputLine in $CommandOutput) {
        Write-Host $OutputLine
    }
    return [int]$ExitCode
}

$SourceProgramBootstrap = @'
import importlib.machinery as machinery
import pathlib
import runpy
import sys

class _SourceOnlyLoader(machinery.SourceFileLoader):
    def get_code(self, fullname):
        source_path = self.get_filename(fullname)
        source_bytes = self.get_data(source_path)
        return self.source_to_code(source_bytes, source_path)

class _SourceOnlyFileFinder(machinery.FileFinder):
    def find_spec(self, fullname, target=None):
        leaf_name = fullname.rpartition('.')[2]
        directory = pathlib.Path(self.path)
        candidates = (
            directory / f'{leaf_name}.pyc',
            directory / f'{leaf_name}.pyo',
            directory / leaf_name / '__init__.pyc',
            directory / leaf_name / '__init__.pyo',
        )
        for candidate in candidates:
            if candidate.is_file():
                raise ImportError(
                    f'Refusing sourceless bytecode import: {candidate}'
                )
        return super().find_spec(fullname, target)

program_path = pathlib.Path(sys.argv[1]).resolve()
source_only_path_hook = _SourceOnlyFileFinder.path_hook(
    (_SourceOnlyLoader, machinery.SOURCE_SUFFIXES),
    (machinery.ExtensionFileLoader, machinery.EXTENSION_SUFFIXES),
)
sys.path_hooks.insert(0, source_only_path_hook)
sys.path_importer_cache.clear()
sys.path.insert(0, str(program_path.parent))
sys.argv = [str(program_path), *sys.argv[2:]]
runpy.run_path(str(program_path), run_name='__main__')
'@

function Invoke-IsolatedPythonSourceProgram {
    param(
        [Parameter(Mandatory = $true)][string[]]$Launcher,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $Executable = $Launcher[0]
    $PrefixArguments = @($Launcher | Select-Object -Skip 1)
    $CommandOutput = @(
        & $Executable @PrefixArguments -c $SourceProgramBootstrap @Arguments
    )
    $ExitCode = $LASTEXITCODE
    foreach ($OutputLine in $CommandOutput) {
        Write-Host $OutputLine
    }
    return [int]$ExitCode
}

function Get-BasePythonIdentity {
    param([Parameter(Mandatory = $true)][string[]]$Launcher)

    $Executable = $Launcher[0]
    $PrefixArguments = @($Launcher | Select-Object -Skip 1)
    $IdentityJson = & $Executable @PrefixArguments -c (
        "import json,sys; print(json.dumps({" +
        "'executable':sys.executable,'version':sys.version.split()[0]," +
        "'major_minor':f'{sys.version_info.major}.{sys.version_info.minor}'}))"
    )
    if ($LASTEXITCODE -ne 0) {
        throw "Could not inspect the Python runtime (exit code $LASTEXITCODE)."
    }
    return ($IdentityJson | ConvertFrom-Json)
}

function Get-FileSha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $Stream = [System.IO.File]::OpenRead($Path)
    $Hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $HashBytes = $Hasher.ComputeHash($Stream)
        return (($HashBytes | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $Hasher.Dispose()
        $Stream.Dispose()
    }
}

function Test-PythonEnvironment {
    if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
        return $false
    }
    & $VenvPython -I -B -m pip check
    if ($LASTEXITCODE -ne 0) {
        return $false
    }
    $VerifyScript = (
        "import importlib.metadata as m,pathlib,re,sys;" +
        "canon=lambda x:re.sub(r'[-_.]+','-',x).lower();" +
        "lines=pathlib.Path(sys.argv[1]).read_text(encoding='utf-8').splitlines();" +
        "req=[x.strip() for x in lines if x.strip() and not x.lstrip().startswith('#') and '==' in x];" +
        "pins=[(x.split('==',1)[0].strip(),x.split('==',1)[1].strip().split()[0]) for x in req];" +
        "installed={canon(d.metadata['Name']):d.version for d in m.distributions() if d.metadata['Name']};" +
        "bad=[n+'=='+v+': installed='+installed.get(canon(n),'missing') for n,v in pins if installed.get(canon(n))!=v];" +
        "print('\\n'.join(bad));raise SystemExit(1 if bad else 0)"
    )
    & $VenvPython -I -B -c $VerifyScript $InstallRequirementsPath
    return $LASTEXITCODE -eq 0
}

function Write-EnvironmentState {
    param(
        [Parameter(Mandatory = $true)]$PythonIdentity,
        [Parameter(Mandatory = $true)][string]$RequirementsSha256
    )

    New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
    $Payload = [ordered]@{
        schema_version = 1
        created_at_utc = [DateTime]::UtcNow.ToString("o")
        python_executable = $PythonIdentity.executable
        python_version = $PythonIdentity.version
        python_major_minor = $PythonIdentity.major_minor
        requirements_file = (Split-Path -Leaf $InstallRequirementsPath)
        requirements_sha256 = $RequirementsSha256
    }
    $TemporaryStatePath = Join-Path (
        Split-Path -Parent $EnvironmentStatePath
    ) (".environment-state." + [guid]::NewGuid().ToString("N") + ".tmp")
    try {
        $Payload | ConvertTo-Json | Set-Content -LiteralPath $TemporaryStatePath -Encoding UTF8
        Move-Item -LiteralPath $TemporaryStatePath -Destination $EnvironmentStatePath -Force
    } finally {
        if (Test-Path -LiteralPath $TemporaryStatePath) {
            Remove-Item -LiteralPath $TemporaryStatePath -Force
        }
    }
}

function Rebuild-PythonEnvironment {
    param(
        [Parameter(Mandatory = $true)][string[]]$BasePython,
        [Parameter(Mandatory = $true)]$PythonIdentity,
        [Parameter(Mandatory = $true)][string]$RequirementsSha256
    )

    $ResolvedStateDir = [System.IO.Path]::GetFullPath($StateDir)
    $ResolvedVenvDir = [System.IO.Path]::GetFullPath($VenvDir)
    $StatePrefix = $ResolvedStateDir.TrimEnd("\") + "\"
    if (-not $ResolvedVenvDir.StartsWith(
        $StatePrefix, [System.StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Refusing to rebuild a Python environment outside $ResolvedStateDir."
    }

    New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
    $BackupDir = Join-Path $StateDir (".venv-backup-" + [guid]::NewGuid().ToString("N"))
    $HadExistingEnvironment = Test-Path -LiteralPath $VenvDir
    if ($HadExistingEnvironment) {
        Move-Item -LiteralPath $VenvDir -Destination $BackupDir
    }
    try {
        $CreateExitCode = Invoke-PythonLauncher -Launcher $BasePython -Arguments @(
            "-m", "venv", $VenvDir
        )
        if ($CreateExitCode -ne 0) {
            throw "Could not create the Red Onion Python environment (exit code $CreateExitCode)."
        }

        $InstallArguments = @(
            "-I", "-B", "-m", "pip", "install", "--disable-pip-version-check", "--quiet"
        )
        if ($InstallRequirementsPath -eq $LockedRequirementsPath) {
            $InstallArguments += "--require-hashes"
        }
        $InstallArguments += @("-r", $InstallRequirementsPath)
        & $VenvPython @InstallArguments
        if ($LASTEXITCODE -ne 0) {
            throw "Could not install the Red Onion program requirements (exit code $LASTEXITCODE)."
        }
        if (-not (Test-PythonEnvironment)) {
            throw "The rebuilt Red Onion Python environment failed dependency verification."
        }
        Write-EnvironmentState -PythonIdentity $PythonIdentity `
            -RequirementsSha256 $RequirementsSha256
        if (Test-Path -LiteralPath $BackupDir) {
            Remove-Item -LiteralPath $BackupDir -Recurse -Force
        }
    } catch {
        if (Test-Path -LiteralPath $VenvDir) {
            Remove-Item -LiteralPath $VenvDir -Recurse -Force
        }
        if (Test-Path -LiteralPath $BackupDir) {
            Move-Item -LiteralPath $BackupDir -Destination $VenvDir
        }
        throw
    }
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

$BasePython = Get-PythonLauncher
$PythonIdentity = Get-BasePythonIdentity -Launcher $BasePython

# Use the standard-library-only validator before creating workflow locks,
# the virtual environment, or package state. The fallback keeps
# older standalone test/copy layouts operable; released copies include the
# dedicated validator.
$ValidationTarget = if (Test-Path -LiteralPath $ConfigValidatorPath) {
    $ConfigValidatorPath
} else {
    $ProgramPath
}
$ValidationArguments = @($ValidationTarget, "--config", $ConfigPath)
if ($ValidationTarget -eq $ProgramPath) {
    $ValidationArguments += "--validate-config"
}
$ValidationExitCode = Invoke-IsolatedPythonSourceProgram -Launcher $BasePython `
    -Arguments $ValidationArguments
if ($ValidationExitCode -ne 0) {
    throw "Configuration validation failed (exit code $ValidationExitCode). No runtime folders or reports were changed."
}

$MaintenanceOnly = $HealthCheck -or [bool]$RebindRestoredIntegrityAnchor
if (-not $MaintenanceOnly) {
    foreach ($dir in @($InputDir, $OutputDir, $ArchiveDir)) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
            throw (
                "Reporting workspace folder unavailable: $dir. " +
                "Download or restore the complete workspace before reporting. " +
                "For a new workspace without reporting history, create its three operator folders once as described in SETUP.md. " +
                "No reporting folders or reports were changed."
            )
        }
    }
}

$RequirementsSha256 = Get-FileSha256 -Path $InstallRequirementsPath
$EnvironmentMatches = $false
if (
    -not $RebuildEnvironment `
    -and (Test-Path -LiteralPath $VenvPython -PathType Leaf) `
    -and (Test-Path -LiteralPath $EnvironmentStatePath -PathType Leaf)
) {
    try {
        $EnvironmentState = Get-Content -LiteralPath $EnvironmentStatePath -Raw |
            ConvertFrom-Json
        $EnvironmentMatches = (
            $EnvironmentState.schema_version -eq 1 `
            -and $EnvironmentState.python_version -eq $PythonIdentity.version `
            -and $EnvironmentState.requirements_file -eq (
                Split-Path -Leaf $InstallRequirementsPath
            ) `
            -and $EnvironmentState.requirements_sha256 -eq $RequirementsSha256
        )
    } catch {
        $EnvironmentMatches = $false
    }
}

if ($HealthCheck -and -not $RebuildEnvironment -and -not $EnvironmentMatches) {
    throw (
        "Health check found that the local Python environment is missing or stale. " +
        "A technical maintainer must run this launcher once with " +
        "-RebuildEnvironment -HealthCheck. " +
        "The health check did not create or install anything."
    )
}

if ($RebuildEnvironment -or (-not $HealthCheck -and -not $EnvironmentMatches)) {
    Write-Host "Building the verified Red Onion Python environment..."
    Rebuild-PythonEnvironment -BasePython $BasePython `
        -PythonIdentity $PythonIdentity -RequirementsSha256 $RequirementsSha256
    $EnvironmentMatches = $true
}
if (-not (Test-PythonEnvironment)) {
    throw (
        "The Red Onion Python environment failed local dependency verification. " +
        "A technical maintainer should rerun with -RebuildEnvironment."
    )
}

$ProgramArguments = @(
    $ProgramPath,
    "--input-dir", $InputDir,
    "--output-dir", $OutputDir,
    "--archive-dir", $ArchiveDir,
    "--config", $ConfigPath
)
# Machine-local anchors belong only to explicit legacy maintenance operations.
if ($InitializeIntegrityBaseline -or $RebindRestoredIntegrityAnchor) {
    $ProgramArguments += @("--integrity-anchor-dir", $IntegrityAnchorDir)
}
if ($InitializeIntegrityBaseline) {
    $ProgramArguments += "--initialize-integrity-baseline"
}
if ($ReviewCorrections) {
    $ProgramArguments += "--review-corrections"
}
if ($CorrectionsFile) {
    $ProgramArguments += @("--corrections-file", $CorrectionsFile)
}
if ($ManagementWorkbook) {
    $ProgramArguments += @("--management-workbook", $ManagementWorkbook)
}
if ($HealthCheck) {
    $ProgramArguments += "--health-check"
}
if ($RebindRestoredIntegrityAnchor) {
    $ProgramArguments += @(
        "--rebind-restored-integrity-anchor",
        $RebindRestoredIntegrityAnchor
    )
}

if ($ReviewCorrections) {
    # A real console must see the question before Python reads the answer.
    # Capturing stdout in the helper would withhold the prompt until process exit.
    & $VenvPython -I -B -c $SourceProgramBootstrap @ProgramArguments
    exit $LASTEXITCODE
}

$ReportExitCode = Invoke-IsolatedPythonSourceProgram `
    -Launcher @($VenvPython, "-I", "-B") -Arguments $ProgramArguments
if ($ReportExitCode -ne 0) {
    exit $ReportExitCode
}

exit 0
