from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import py_compile
import queue
import shutil
import subprocess
import sys
import threading
from pathlib import Path

import pytest


PROGRAM_DIR = Path(__file__).resolve().parents[1]
LAUNCHER = PROGRAM_DIR / "Run-WeeklySnapshot.ps1"
POWERSHELL = shutil.which("powershell.exe")
GIT = shutil.which("git")
WINDOWS_LAUNCHER_AVAILABLE = os.name == "nt" and POWERSHELL is not None


def test_launcher_retains_isolated_source_loading_and_prewrite_config_validation() -> None:
    launcher = LAUNCHER.read_text(encoding="utf-8")
    assert "Invoke-IsolatedPythonSourceProgram" in launcher
    assert "class _SourceOnlyLoader" in launcher
    assert 'return @("py", $PythonSelector, "-I", "-B")' in launcher
    assert 'return @("python", "-I", "-B")' in launcher
    assert '--require-hashes' in launcher
    assert launcher.index(
        "$ValidationExitCode = Invoke-IsolatedPythonSourceProgram"
    ) < launcher.index("$RequirementsSha256 = Get-FileSha256")
    assert "Assert-DeployedRelease" not in launcher
    assert "Assert-NoSourceBytecode" not in launcher


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
    # Minimal test programs use an empty requirements file. Record that already
    # installed environment so each module does not create the same venv twice.
    (venv_dir.parent / "environment-state.json").write_text(
        json.dumps({
            "schema_version": 1,
            "python_version": sys.version.split()[0],
            "requirements_file": "requirements.txt",
            "requirements_sha256": hashlib.sha256(b"").hexdigest(),
        }), encoding="utf-8",
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
def test_git_checkout_remains_supported(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    _initialize_release_checkout(repository_root)

    result = _run_launcher(launcher, launcher_environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Verified deployed release" not in result.stdout
    assert (repository_root / "01 Daily Reports - Drop Here").is_dir()


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
def test_source_bootstrap_rejects_sourceless_import_after_config_validation(
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
def test_unused_cache_and_harmless_notes_do_not_block_reporting(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    _initialize_release_checkout(repository_root)
    (repository_root / "local-note.txt").write_text("operator note", encoding="utf-8")
    cache = launcher.parent / "__pycache__"
    cache.mkdir()
    (cache / "stale_module.pyc").write_bytes(b"unused old cache")
    (launcher.parent / "unused.pyo").write_bytes(b"unused old optimized cache")

    result = _run_launcher(launcher, launcher_environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert (repository_root / "01 Daily Reports - Drop Here").is_dir()


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


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_named_deployment_without_git_runs_without_machine_anchor(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    capture_path = tmp_path / "arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)

    result = _run_launcher(launcher, environment)

    assert result.returncode == 0, result.stdout + result.stderr
    assert "--integrity-anchor-dir" not in capture_path.read_text(encoding="utf-8")
    assert (repository_root / "01 Daily Reports - Drop Here").is_dir()


@pytest.mark.skipif(
    not WINDOWS_LAUNCHER_AVAILABLE or GIT is None,
    reason="Windows PowerShell and Git required",
)
@pytest.mark.parametrize(
    "invalid_state",
    ["dirty", "wrong-branch", "detached", "missing-origin", "head-mismatch"],
)
def test_git_worktree_state_does_not_gate_reporting(
    tmp_path: Path,
    launcher_environment: dict[str, str],
    invalid_state: str,
) -> None:
    repository_root = tmp_path / "Red Onion Weekly Metrics Automation"
    launcher = _write_minimal_runner(repository_root)
    _initialize_release_checkout(repository_root)

    if invalid_state == "dirty":
        (repository_root / "local-note.txt").write_text("unreleased", encoding="utf-8")
    elif invalid_state == "wrong-branch":
        _git(repository_root, "checkout", "-b", "manager-edit")
    elif invalid_state == "detached":
        _git(repository_root, "checkout", "--detach")
    elif invalid_state == "missing-origin":
        _git(repository_root, "update-ref", "-d", "refs/remotes/origin/main")
    else:
        (repository_root / "new-release.txt").write_text("ahead", encoding="utf-8")
        _git(repository_root, "add", "new-release.txt")
        _git(repository_root, "commit", "-m", "unpublished release")

    result = _run_launcher(launcher, launcher_environment)
    combined_output = result.stdout + result.stderr
    normalized_output = " ".join(combined_output.split())

    assert result.returncode == 0, combined_output
    assert "Release preflight failed" not in normalized_output
    assert (repository_root / "01 Daily Reports - Drop Here").is_dir()


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
@pytest.mark.parametrize("layout", ["own", "parent", "renamed", "explicit"])
def test_operator_root_resolves_by_folder_evidence_or_explicit_argument(
    tmp_path: Path, launcher_environment: dict[str, str], layout: str
) -> None:
    repository_root = tmp_path / "Downloaded Program"
    launcher = _write_minimal_runner(repository_root)
    expected = repository_root if layout == "own" else tmp_path
    (expected / "01 Daily Reports - Drop Here").mkdir(exist_ok=True)
    if layout == "renamed":
        renamed = tmp_path / "Any program name"
        repository_root.rename(renamed)
        repository_root = renamed
        launcher = renamed / "_program" / LAUNCHER.name
    arguments: list[str] = []
    if layout == "explicit":
        (repository_root / "03 Archive").mkdir()
        expected = tmp_path / "Chosen workspace"
        arguments = ["-OperationsRoot", str(expected)]
    capture_path = tmp_path / "root-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)

    result = _run_launcher(launcher, environment, *arguments)

    assert result.returncode == 0, result.stdout + result.stderr
    forwarded = capture_path.read_text(encoding="utf-8").splitlines()
    for option, folder in [("--input-dir", "01 Daily Reports - Drop Here"),
                           ("--output-dir", "02 Finished Reports"),
                           ("--archive-dir", "03 Archive")]:
        assert Path(forwarded[forwarded.index(option) + 1]) == expected / folder


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
@pytest.mark.parametrize("layout", ["standalone", "renamed-child", "explicit"])
def test_cmd_noninteractive_routes_without_pause_or_false_recovery_claim(
    tmp_path: Path, layout: str
) -> None:
    root_cmd = tmp_path / "Run Weekly Snapshot.cmd"
    shutil.copy2(PROGRAM_DIR.parent / root_cmd.name, root_cmd)
    program_root = tmp_path if layout == "standalone" else tmp_path / "Renamed program"
    runner = program_root / "_program" / LAUNCHER.name
    runner.parent.mkdir(parents=True)
    runner.write_text(
        'param([string]$OperationsRoot, [switch]$NonInteractive)\n'
        'Write-Output "ROOT=$OperationsRoot"\n'
        'Write-Output "Published with cleanup warning"\n'
        'exit 7\n', encoding="utf-8"
    )
    arguments = ["-NonInteractive"]
    if layout == "explicit":
        arguments += ["-OperationsRoot", str(tmp_path / "Explicit workspace")]
    result = subprocess.run(
        [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", "call", str(root_cmd), *arguments],
        cwd=tmp_path, capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=30,
    )
    assert result.returncode == 7, result.stdout + result.stderr
    assert "Published with cleanup warning" in result.stdout
    assert "Press any key" not in result.stdout
    assert "reports were left in place" not in result.stdout
    if layout == "renamed-child":
        assert f"ROOT={tmp_path}" in result.stdout
    elif layout == "explicit":
        assert f"ROOT={tmp_path / 'Explicit workspace'}" in result.stdout


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_cmd_refuses_to_guess_between_two_program_copies(tmp_path: Path) -> None:
    root_cmd = tmp_path / "Run Weekly Snapshot.cmd"
    shutil.copy2(PROGRAM_DIR.parent / root_cmd.name, root_cmd)
    for name in ["Current program", "Other program"]:
        runner = tmp_path / name / "_program" / LAUNCHER.name
        runner.parent.mkdir(parents=True)
        runner.write_text('throw "must not run"', encoding="utf-8")
    result = subprocess.run(
        [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", "call", str(root_cmd), "-NonInteractive"],
        cwd=tmp_path, capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=30,
    )
    assert result.returncode == 2
    assert "Expected one program folder" in result.stdout
    assert "must not run" not in result.stderr
    assert "Press any key" not in result.stdout


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_explicit_correction_and_master_choices_are_forwarded(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    launcher = _write_minimal_runner(tmp_path / "program")
    capture_path = tmp_path / "choice-arguments.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    corrections = str(tmp_path / "Reviewed choices.json")
    master = str(tmp_path / "Chosen master.xlsx")

    result = _run_launcher(
        launcher, environment, "-CorrectionsFile", corrections,
        "-ManagementWorkbook", master, "-NonInteractive",
    )

    assert result.returncode == 0, result.stdout + result.stderr
    forwarded = capture_path.read_text(encoding="utf-8").splitlines()
    assert forwarded[forwarded.index("--corrections-file") + 1] == corrections
    assert forwarded[forwarded.index("--management-workbook") + 1] == master
    assert "--review-corrections" not in forwarded

    result = _run_launcher(launcher, environment, "-ReviewCorrections")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "--review-corrections" in capture_path.read_text(encoding="utf-8")


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_noninteractive_review_is_rejected_before_python(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    launcher = _write_minimal_runner(tmp_path / "program")
    capture_path = tmp_path / "must-not-exist.txt"
    environment = launcher_environment.copy()
    environment["RED_ONION_TEST_ARGUMENTS_PATH"] = str(capture_path)
    result = _run_launcher(launcher, environment, "-NonInteractive", "-ReviewCorrections")
    assert result.returncode != 0
    assert "needs an interactive review" in result.stdout + result.stderr
    assert not capture_path.exists()
    assert not (launcher.parent.parent / "03 Archive").exists()


@pytest.mark.skipif(not WINDOWS_LAUNCHER_AVAILABLE, reason="Windows PowerShell required")
def test_review_question_is_visible_before_an_answer_is_supplied(
    tmp_path: Path, launcher_environment: dict[str, str]
) -> None:
    launcher = _write_minimal_runner(tmp_path / "program")
    (launcher.parent / "red_onion_weekly_metrics.py").write_text(
        'import sys\n'
        'if "--review-corrections" in sys.argv:\n'
        '    print("Choose correction:", flush=True)\n'
        '    print("Chosen=" + input())\n', encoding="utf-8",
    )
    # Finish environment setup before starting the prompt visibility deadline.
    setup = _run_launcher(launcher, launcher_environment)
    assert setup.returncode == 0, setup.stdout + setup.stderr
    process = subprocess.Popen(
        [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
         str(launcher), "-ReviewCorrections"],
        cwd=launcher.parent.parent, env=launcher_environment,
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
    )
    visible: queue.Queue[str] = queue.Queue()
    observed: list[str] = []

    def read_question() -> None:
        text = ""
        assert process.stdout is not None
        while True:
            character = process.stdout.read(1)
            text += character
            observed.append(character)
            if not character or "Choose correction:" in text:
                visible.put(text)
                return

    reader = threading.Thread(target=read_question, daemon=True)
    reader.start()
    try:
        try:
            output = visible.get(timeout=90)
        except queue.Empty:
            pytest.fail("Question did not become visible: " + "".join(observed))
        assert "Choose correction:" in output, output
        assert process.stdin is not None
        process.stdin.write("2\n")
        process.stdin.flush()
        remaining, _ = process.communicate(timeout=30)
        assert process.returncode == 0, output + remaining
        assert "Chosen=2" in remaining
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=10)
        reader.join(timeout=10)
