from __future__ import annotations

import os
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
    preflight_call = launcher.index(
        "if ($IsDeployedCheckout) {\n    Assert-DeployedRelease"
    )
    assert preflight_call < launcher.index(
        "New-Item -ItemType Directory"
    )
    assert preflight_call < launcher.index("-m pip install")
    assert "[switch]$InitializeIntegrityBaseline" in launcher
    assert '$ProgramArguments += "--initialize-integrity-baseline"' in launcher


def test_production_dependency_pins_are_exact_and_consistent() -> None:
    expected = ["pandas==2.3.3", "openpyxl==3.1.5", "xlrd==2.0.2"]
    requirements = (PROGRAM_DIR / "requirements.txt").read_text(
        encoding="utf-8"
    ).splitlines()
    pyproject = (PROGRAM_DIR / "pyproject.toml").read_text(encoding="utf-8")

    assert requirements == expected
    for dependency in expected:
        assert f'"{dependency}"' in pyproject
    assert 'py-modules = ["red_onion_integrity", "red_onion_weekly_metrics"]' in pyproject


def _write_minimal_runner(repository_root: Path) -> Path:
    program_dir = repository_root / "_program"
    program_dir.mkdir(parents=True)
    shutil.copy2(LAUNCHER, program_dir / LAUNCHER.name)
    (program_dir / "requirements.txt").write_text("", encoding="utf-8")
    (program_dir / "red_onion_weekly_metrics.py").write_text(
        "raise SystemExit(0)\n", encoding="utf-8"
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


def _run_launcher(launcher: Path, environment: dict[str, str]) -> subprocess.CompletedProcess[str]:
    assert POWERSHELL is not None
    return subprocess.run(
        [
            POWERSHELL,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(launcher),
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
