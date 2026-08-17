from __future__ import annotations

import importlib.util
import os
import py_compile
import shutil
import subprocess
import sys
from pathlib import Path

import pytest


PROGRAM_DIR = Path(__file__).resolve().parents[1]
LAUNCHER = PROGRAM_DIR / "Run-WeeklySnapshot.ps1"
POWERSHELL = shutil.which("powershell.exe")
GIT = shutil.which("git")
WINDOWS_LAUNCHER_AVAILABLE = os.name == "nt" and POWERSHELL is not None


def test_release_preflight_runs_before_any_runtime_mutation() -> None:
    launcher = LAUNCHER.read_text(encoding="utf-8")
    compact_launcher = " ".join(launcher.split())

    assert "$IsDeployedCheckout" in launcher
    assert '"status", "--porcelain=v1", "--untracked-files=all"' in compact_launcher
    assert '"symbolic-ref", "--quiet", "--short", "HEAD"' in compact_launcher
    assert 'refs/remotes/origin/main^{commit}' in launcher
    assert "function Assert-NoSourceBytecode" in launcher
    assert 'Where-Object { $_.Extension -in @(".pyc", ".pyo") }' in launcher
    assert "-ErrorAction Stop" in launcher
    assert "could not be completely inspected" in launcher
    assert 'PYTHONDONTWRITEBYTECODE = "1"' in launcher
    assert "PYTHONPYCACHEPREFIX" in launcher
    assert "Invoke-IsolatedPythonSourceProgram" in launcher
    assert "Refusing sourceless bytecode import" in launcher
    assert "class _SourceOnlyLoader" in launcher
    assert "self.get_data(source_path)" in launcher
    assert "sys.path_hooks.insert(0, source_only_path_hook)" in launcher
    assert 'return @("py", $PythonSelector, "-I", "-B")' in launcher
    assert 'return @("python", "-I", "-B")' in launcher
    assert '& $VenvPython -I -B -m pip check' in launcher
    preflight_call = launcher.index(
        "if ($IsDeployedCheckout) {\n    Assert-DeployedRelease"
    )
    bytecode_call = launcher.index(
        "    Assert-NoSourceBytecode", preflight_call
    )
    assert preflight_call < bytecode_call
    assert bytecode_call < launcher.index('PYTHONDONTWRITEBYTECODE = "1"')
    assert preflight_call < launcher.index(
        "New-Item -ItemType Directory"
    )
    assert preflight_call < launcher.index("$InstallArguments = @(")
    assert "[switch]$InitializeIntegrityBaseline" in launcher
    assert '$ProgramArguments += "--initialize-integrity-baseline"' in launcher
    assert "[switch]$RebuildEnvironment" in launcher
    assert "[switch]$HealthCheck" in launcher
    assert "[string]$RebindRestoredIntegrityAnchor" in launcher
    assert "environment-state.json" in launcher
    assert "red_onion_config.py" in launcher
    config_preflight = launcher.index(
        "$ValidationExitCode = Invoke-IsolatedPythonSourceProgram"
    )
    assert config_preflight < launcher.index(
        "foreach ($dir in @($InputDir, $OutputDir, $ArchiveDir))"
    )
    assert (
        "if ($RebuildEnvironment -or (-not $HealthCheck "
        "-and -not $EnvironmentMatches))" in compact_launcher
    )
    assert "--require-hashes" in launcher
    assert '@("-3.12", "-3.11", "-3.10")' in launcher
    assert launcher.count("(3, 10) <= sys.version_info[:2] < (3, 13)") == 2
    assert "Python 3.10-3.12 was not found" in launcher


def test_environment_reuse_verifies_every_locked_dependency_pin() -> None:
    launcher = LAUNCHER.read_text(encoding="utf-8")
    compact_launcher = " ".join(launcher.split())

    assert "$InstallRequirementsPath" in launcher
    assert (
        "$VenvPython -I -B -c $VerifyScript $InstallRequirementsPath"
        in compact_launcher
    )
    assert "x.split('==',1)[1].strip().split()[0]" in launcher
    assert "m.distributions()" in launcher
    assert (
        "$VenvPython -I -B -c $VerifyScript $DirectRequirementsPath"
        not in compact_launcher
    )


def test_production_dependency_pins_are_exact_and_consistent() -> None:
    expected = ["pandas==2.3.3", "openpyxl==3.1.5", "xlrd==2.0.2"]
    requirements = (PROGRAM_DIR / "requirements.txt").read_text(
        encoding="utf-8"
    ).splitlines()
    pyproject = (PROGRAM_DIR / "pyproject.toml").read_text(encoding="utf-8")

    assert requirements == expected
    for dependency in expected:
        assert f'"{dependency}"' in pyproject
    for module in (
        "red_onion_config",
        "red_onion_integrity",
        "red_onion_runtime",
        "red_onion_weekly_metrics",
    ):
        assert f'"{module}"' in pyproject
    assert 'requires-python = ">=3.10,<3.13"' in pyproject
    lock = (PROGRAM_DIR / "requirements.lock").read_text(encoding="utf-8")
    assert "--hash=sha256:" in lock
    assert "numpy==2.2.6" in lock
    for dependency in expected:
        assert dependency in lock


def _write_minimal_runner(repository_root: Path) -> Path:
    program_dir = repository_root / "_program"
    program_dir.mkdir(parents=True)
    shutil.copy2(LAUNCHER, program_dir / LAUNCHER.name)
    (program_dir / "requirements.txt").write_text("", encoding="utf-8")
    (program_dir / "red_onion_weekly_metrics.py").write_text(
        """from __future__ import annotations

import os
import sys
from pathlib import Path

capture_path = os.environ.get("RED_ONION_TEST_ARGUMENTS_PATH")
if capture_path:
    Path(capture_path).write_text("\\n".join(sys.argv[1:]), encoding="utf-8")
environment_path = os.environ.get("RED_ONION_TEST_ENVIRONMENT_PATH")
if environment_path:
    Path(environment_path).write_text(
        "\\n".join(
            [
                os.environ.get("PYTHONDONTWRITEBYTECODE", ""),
                os.environ.get("PYTHONPYCACHEPREFIX", ""),
            ]
        ),
        encoding="utf-8",
    )
raise SystemExit(0)
""",
        encoding="utf-8",
    )
    return program_dir / LAUNCHER.name


def _git(repository_root: Path, *arguments: str) -> subprocess.CompletedProcess[str]:
    assert GIT is not None
    return subprocess.run(
        [GIT, *arguments],
        cwd=repository_root,
        check=True,
        capture_output=True,
        text=True,
    )


def _initialize_release_checkout(repository_root: Path) -> None:
    _git(repository_root, "init")
    _git(repository_root, "config", "user.email", "launcher-test@example.invalid")
    _git(repository_root, "config", "user.name", "Launcher Test")
    _git(repository_root, "checkout", "-B", "main")
    _git(repository_root, "add", ".")
    _git(repository_root, "commit", "-m", "test release")
    _git(repository_root, "update-ref", "refs/remotes/origin/main", "HEAD")


def _run_launcher(
    launcher: Path,
    environment: dict[str, str],
    *launcher_arguments: str,
) -> subprocess.CompletedProcess[str]:
    assert POWERSHELL is not None
    return subprocess.run(
        [
            POWERSHELL,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(launcher),
            *launcher_arguments,
        ],
        cwd=launcher.parent.parent,
        env=environment,
        capture_output=True,
        text=True,
    )


@pytest.fixture(scope="module")
def launcher_environment(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    if not WINDOWS_LAUNCHER_AVAILABLE:
        pytest.skip("PowerShell launcher behavior is Windows-specific")

    local_app_data = tmp_path_factory.mktemp("launcher-local-app-data")
    venv_dir = local_app_data / "RedOnionMetrics" / ".venv"
    subprocess.run(
        [sys.executable, "-m", "venv", str(venv_dir)],
        check=True,
        capture_output=True,
        text=True,
    )
    environment = os.environ.copy()
    environment["LOCALAPPDATA"] = str(local_app_data)
    return environment


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_standalone_non_git_copy_remains_supported(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "standalone-copy"
    launcher = _write_minimal_runner(repository_root)

    result = _run_launcher(launcher, launcher_environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Verified deployed release" not in result.stdout
    assert (repository_root / "01 Daily Reports - Drop Here").is_dir()
    assert (repository_root / "02 Finished Reports").is_dir()
    assert (repository_root / "03 Archive").is_dir()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_clean_main_deployment_at_local_origin_main_runs(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    _initialize_release_checkout(repository_root)

    result = _run_launcher(launcher, launcher_environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Verified deployed release: main at" in result.stdout
    assert (tmp_path / "01 Daily Reports - Drop Here").is_dir()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_deployed_checkout_isolates_caller_working_directory_bytecode(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    (repository_root / ".gitignore").write_text("*.pyc\n", encoding="utf-8")
    _initialize_release_checkout(repository_root)

    execution_marker = tmp_path / "working-directory-bytecode-executed.txt"
    malicious_source = tmp_path / "attacker_json.py"
    malicious_source.write_text(
        "from pathlib import Path\n"
        f"Path({str(execution_marker)!r}).write_text("
        "'executed', encoding='utf-8')\n",
        encoding="utf-8",
    )
    py_compile.compile(
        str(malicious_source),
        cfile=str(repository_root / "json.pyc"),
        doraise=True,
    )
    assert _git(repository_root, "status", "--porcelain=v1").stdout == ""

    capture_path = tmp_path / "isolated-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    result = _run_launcher(launcher, environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert capture_path.exists()
    assert not execution_marker.exists()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_source_bootstrap_rejects_bytecode_added_after_initial_preflight(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    program_path = launcher.parent / "red_onion_weekly_metrics.py"
    program_path.write_text(
        """from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

if "--validate-config" in sys.argv:
    shutil.copy2(
        os.environ["RED_ONION_LATE_BYTECODE"],
        Path(__file__).with_name("pandas.pyc"),
    )
    raise SystemExit(0)

import pandas

capture_path = os.environ.get("RED_ONION_TEST_ARGUMENTS_PATH")
if capture_path:
    Path(capture_path).write_text("invoked", encoding="utf-8")
raise SystemExit(0)
""",
        encoding="utf-8",
    )
    (repository_root / ".gitignore").write_text("*.pyc\n", encoding="utf-8")
    _initialize_release_checkout(repository_root)

    execution_marker = tmp_path / "late-bytecode-executed.txt"
    malicious_source = tmp_path / "late_attacker_pandas.py"
    malicious_source.write_text(
        "from pathlib import Path\n"
        f"Path({str(execution_marker)!r}).write_text("
        "'executed', encoding='utf-8')\n",
        encoding="utf-8",
    )
    late_bytecode = tmp_path / "late-pandas.pyc"
    py_compile.compile(
        str(malicious_source), cfile=str(late_bytecode), doraise=True
    )

    capture_path = tmp_path / "late-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_LATE_BYTECODE"] = str(late_bytecode)
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    result = _run_launcher(launcher, environment)
    normalized_output = " ".join((result.stdout + result.stderr).split())

    assert result.returncode != 0
    assert "Refusing sourceless bytecode import" in normalized_output
    assert not capture_path.exists()
    assert not execution_marker.exists()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_source_bootstrap_ignores_late_source_backed_cache(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    helper_path = launcher.parent / "trusted_helper.py"
    helper_path.write_text("VALUE = 'source'\n", encoding="utf-8")
    program_path = launcher.parent / "red_onion_weekly_metrics.py"
    program_path.write_text(
        """from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

if "--validate-config" in sys.argv:
    cache_destination = Path(os.environ["RED_ONION_CACHE_DESTINATION"])
    cache_destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(
        os.environ["RED_ONION_LATE_BYTECODE"], cache_destination
    )
    raise SystemExit(0)

import trusted_helper

capture_path = os.environ.get("RED_ONION_TEST_ARGUMENTS_PATH")
if capture_path:
    Path(capture_path).write_text(trusted_helper.VALUE, encoding="utf-8")
raise SystemExit(0)
""",
        encoding="utf-8",
    )
    (repository_root / ".gitignore").write_text(
        "__pycache__/\n*.pyc\n", encoding="utf-8"
    )
    _initialize_release_checkout(repository_root)

    execution_marker = tmp_path / "source-cache-executed.txt"
    malicious_source = tmp_path / "attacker_trusted_helper.py"
    malicious_source.write_text(
        "from pathlib import Path\n"
        f"Path({str(execution_marker)!r}).write_text("
        "'executed', encoding='utf-8')\n"
        "VALUE = 'bytecode'\n",
        encoding="utf-8",
    )
    late_bytecode = tmp_path / "trusted-helper-cache.pyc"
    py_compile.compile(
        str(malicious_source),
        cfile=str(late_bytecode),
        doraise=True,
        invalidation_mode=py_compile.PycInvalidationMode.UNCHECKED_HASH,
    )
    cache_destination = Path(importlib.util.cache_from_source(str(helper_path)))

    capture_path = tmp_path / "source-cache-result.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_LATE_BYTECODE"] = str(late_bytecode)
    environment["RED_ONION_CACHE_DESTINATION"] = str(cache_destination)
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    result = _run_launcher(launcher, environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert capture_path.read_text(encoding="utf-8") == "source"
    assert cache_destination.exists()
    assert not execution_marker.exists()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_deployed_checkout_rejects_ignored_sourceless_bytecode_before_python(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    program_path = launcher.parent / "red_onion_weekly_metrics.py"
    program_path.write_text(
        "import pandas\n" + program_path.read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    (repository_root / ".gitignore").write_text(
        "__pycache__/\n*.pyc\n*.pyo\n", encoding="utf-8"
    )
    _initialize_release_checkout(repository_root)

    execution_marker = tmp_path / "bytecode-executed.txt"
    malicious_source = tmp_path / "attacker_pandas.py"
    malicious_source.write_text(
        "from pathlib import Path\n"
        f"Path({str(execution_marker)!r}).write_text("
        "'executed', encoding='utf-8')\n",
        encoding="utf-8",
    )
    bytecode = launcher.parent / "pandas.pyc"
    py_compile.compile(
        str(malicious_source),
        cfile=str(bytecode),
        doraise=True,
    )
    optimized_bytecode = launcher.parent / "stale_module.pyo"
    optimized_bytecode.write_bytes(b"unverified optimized bytecode")
    assert _git(repository_root, "status", "--porcelain=v1").stdout == ""

    capture_path = tmp_path / "invoked-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    result = _run_launcher(launcher, environment)
    normalized_output = " ".join((result.stdout + result.stderr).split())

    assert result.returncode != 0
    assert "Release preflight failed" in normalized_output
    assert "Python bytecode that cannot be verified by Git" in normalized_output
    assert "_program\\pandas.pyc" in normalized_output
    assert "_program\\stale_module.pyo" in normalized_output
    assert not capture_path.exists()
    assert not execution_marker.exists()
    assert not (tmp_path / "01 Daily Reports - Drop Here").exists()


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_initialize_integrity_baseline_argument_is_forwarded(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "standalone-baseline"
    launcher = _write_minimal_runner(repository_root)
    capture_path = tmp_path / "baseline-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)

    result = _run_launcher(
        launcher,
        environment,
        "-InitializeIntegrityBaseline",
    )

    assert result.returncode == 0, result.stdout + result.stderr
    forwarded_arguments = capture_path.read_text(encoding="utf-8").splitlines()
    assert "--initialize-integrity-baseline" in forwarded_arguments
    anchor_flag_index = forwarded_arguments.index("--integrity-anchor-dir")
    assert forwarded_arguments[anchor_flag_index + 1] == str(
        Path(environment["LOCALAPPDATA"])
        / "RedOnionMetrics"
        / "integrity-anchors"
    )
    assert not list(launcher.parent.rglob("*.pyc"))


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_health_check_does_not_create_operator_folders(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "standalone-health"
    launcher = _write_minimal_runner(repository_root)
    environment = launcher_environment.copy()
    capture_path = tmp_path / "forwarded-arguments.txt"
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    initial = _run_launcher(launcher, environment)
    assert initial.returncode == 0, initial.stdout + initial.stderr
    for name in (
        "01 Daily Reports - Drop Here",
        "02 Finished Reports",
        "03 Archive",
    ):
        shutil.rmtree(repository_root / name)

    result = _run_launcher(launcher, environment, "-HealthCheck")

    assert result.returncode == 0, result.stdout + result.stderr
    for name in (
        "01 Daily Reports - Drop Here",
        "02 Finished Reports",
        "03 Archive",
    ):
        assert not (repository_root / name).exists()

    restored_anchor = tmp_path / "backed-up-anchor.json"
    restored_anchor.write_text("{}", encoding="utf-8")
    rebind_result = _run_launcher(
        launcher,
        environment,
        "-RebindRestoredIntegrityAnchor",
        str(restored_anchor),
    )

    assert rebind_result.returncode == 0, (
        rebind_result.stdout + rebind_result.stderr
    )
    forwarded_arguments = capture_path.read_text(encoding="utf-8").splitlines()
    rebind_index = forwarded_arguments.index(
        "--rebind-restored-integrity-anchor"
    )
    assert forwarded_arguments[rebind_index + 1] == str(restored_anchor)
    for name in (
        "01 Daily Reports - Drop Here",
        "02 Finished Reports",
        "03 Archive",
    ):
        assert not (repository_root / name).exists()

    rebuild_result = _run_launcher(
        launcher,
        environment,
        "-RebuildEnvironment",
        "-HealthCheck",
    )

    assert rebuild_result.returncode == 0, (
        rebuild_result.stdout + rebuild_result.stderr
    )
    for name in (
        "01 Daily Reports - Drop Here",
        "02 Finished Reports",
        "03 Archive",
    ):
        assert not (repository_root / name).exists()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_named_deployment_without_git_fails_closed(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)

    result = _run_launcher(launcher, launcher_environment)
    combined_output = result.stdout + result.stderr
    normalized_output = " ".join(combined_output.split())

    assert result.returncode != 0
    assert "Release preflight failed" in normalized_output
    assert "not a Git checkout" in normalized_output
    assert not (tmp_path / "01 Daily Reports - Drop Here").exists()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
@pytest.mark.parametrize(
    "invalid_state",
    ["dirty", "wrong-branch", "detached", "missing-origin", "head-mismatch"],
)
def test_deployed_checkout_fails_closed_before_runtime_changes(
    tmp_path: Path,
    launcher_environment: dict[str, str],
    invalid_state: str,
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    _initialize_release_checkout(repository_root)

    if invalid_state == "dirty":
        (repository_root / "local-note.txt").write_text("unreleased", encoding="utf-8")
        expected = "local or untracked changes"
    elif invalid_state == "wrong-branch":
        _git(repository_root, "checkout", "-b", "manager-edit")
        expected = "instead of 'main'"
    elif invalid_state == "detached":
        _git(repository_root, "checkout", "--detach")
        expected = "detached HEAD instead of 'main'"
    elif invalid_state == "missing-origin":
        _git(repository_root, "update-ref", "-d", "refs/remotes/origin/main")
        expected = "local origin/main reference is missing"
    else:
        (repository_root / "new-release.txt").write_text("ahead", encoding="utf-8")
        _git(repository_root, "add", "new-release.txt")
        _git(repository_root, "commit", "-m", "unpublished release")
        expected = "does not match local origin/main"

    result = _run_launcher(launcher, launcher_environment)
    combined_output = result.stdout + result.stderr
    normalized_output = " ".join(combined_output.split())

    assert result.returncode != 0
    assert "Release preflight failed" in normalized_output
    assert expected in normalized_output
    assert "No reports were created and no source files were moved" in normalized_output
    assert not (tmp_path / "01 Daily Reports - Drop Here").exists()
    assert not (tmp_path / "02 Finished Reports").exists()
    assert not (tmp_path / "03 Archive").exists()
