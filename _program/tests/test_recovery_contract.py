from __future__ import annotations

import os
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]
POWERSHELL = shutil.which("powershell.exe")
GIT = shutil.which("git")


def test_recovery_bundle_defaults_to_released_source_only() -> None:
    script = (ROOT / "Build-RecoveryBundle.ps1").read_text(encoding="utf-8")

    assert "[switch]$IncludeOperationalData" in script
    assert "if ($IncludeOperationalData)" in script
    assert "status\", \"--porcelain=v1\", \"--untracked-files=all\"" in script
    assert "symbolic-ref\", \"--quiet\", \"--short\", \"HEAD\"" in script
    assert "refs/remotes/origin/main^{commit}" in script
    assert "tag\", \"--points-at\", \"HEAD\", \"--list\", \"v*\"" in script
    assert "released-source.zip" in script
    assert "repository.bundle" in script
    assert "SHA256SUMS.txt" in script
    assert "automatic_upload = $false" in script
    assert "automatic_send = $false" in script
    assert ".weekly-snapshot.lock" in script
    assert "Get-OperationalSourceFingerprint" in script
    assert "Operational source state changed during recovery capture" in script
    assert "Test-UnsafeReparsePoint" in script
    assert "Test-SafeCloudReparsePoint" in script
    assert 'if (-not ("ReparseTag" -as [type]))' in script
    assert '$Entry.PSObject.Properties["LinkType"]' in script
    assert '$Entry.PSObject.Properties["Target"]' in script
    assert "GetFileInformationByHandleEx" in script
    assert "0x9000001A" in script
    assert "NameSurrogateBit" in script


def test_recovery_and_governance_contracts_are_explicit() -> None:
    recovery = (ROOT / "RECOVERY.md").read_text(encoding="utf-8")
    governance = (ROOT / "DATA_GOVERNANCE.md").read_text(encoding="utf-8")
    incident = (ROOT / "INCIDENT_RESPONSE.md").read_text(encoding="utf-8")

    assert "newest 13" in recovery
    assert "newest 12" in recovery
    assert "quarterly" in recovery.casefold()
    assert "365 days" in governance
    assert "exact" in governance.casefold()
    assert "does not silently delete" in governance
    assert "Do not initialize a replacement baseline" in incident


@pytest.mark.skipif(
    os.name != "nt" or POWERSHELL is None or GIT is None,
    reason="Windows PowerShell and Git required",
)
def test_source_only_recovery_bundle_is_restorable_and_excludes_operating_data(
    tmp_path: Path,
) -> None:
    repository = tmp_path / "repository"
    repository.mkdir()
    shutil.copy2(
        ROOT / "Build-RecoveryBundle.ps1",
        repository / "Build-RecoveryBundle.ps1",
    )
    program = repository / "_program"
    program.mkdir()
    (program / "red_onion_config.json").write_text("{}", encoding="utf-8")
    (program / "requirements.txt").write_text("", encoding="utf-8")
    (program / "requirements.lock").write_text("", encoding="utf-8")
    (program / "requirements-constraints.txt").write_text("", encoding="utf-8")
    (program / "pyproject.toml").write_text(
        '[project]\nname="test"\nversion="0.2.0"\n',
        encoding="utf-8",
    )
    (repository / ".gitignore").write_text("*.xlsx\n", encoding="utf-8")
    (repository / "operating-secret.xlsx").write_bytes(b"not tracked")

    def git(*arguments: str) -> None:
        subprocess.run(
            [GIT, *arguments],
            cwd=repository,
            check=True,
            capture_output=True,
            text=True,
        )

    git("init")
    git("config", "user.email", "recovery-test@example.invalid")
    git("config", "user.name", "Recovery Test")
    git("checkout", "-B", "main")
    git("add", "Build-RecoveryBundle.ps1", "_program", ".gitignore")
    git("commit", "-m", "released fixture")
    git("update-ref", "refs/remotes/origin/main", "HEAD")
    git("tag", "-a", "v0.2.0", "-m", "fixture release")

    destination = tmp_path / "recovery"
    result = subprocess.run(
        [
            POWERSHELL,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(repository / "Build-RecoveryBundle.ps1"),
            "-DestinationDirectory",
            str(destination),
        ],
        cwd=repository,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stdout + result.stderr
    bundles = list(destination.glob("*.zip"))
    assert len(bundles) == 1
    assert Path(str(bundles[0]) + ".sha256.txt").is_file()
    with zipfile.ZipFile(bundles[0]) as archive:
        names = set(archive.namelist())
    assert "released-source.zip" in names
    assert "repository.bundle" in names
    assert "release-metadata.json" in names
    assert "SHA256SUMS.txt" in names
    assert all("operating-secret" not in name for name in names)

    blocked = subprocess.run(
        [
            POWERSHELL,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(repository / "Build-RecoveryBundle.ps1"),
            "-DestinationDirectory",
            str(repository),
        ],
        cwd=repository,
        capture_output=True,
        text=True,
    )
    assert blocked.returncode != 0
    assert "outside the Git repository" in (blocked.stdout + blocked.stderr)

    operations_root = tmp_path / "operations"
    required_folders = (
        operations_root / "02 Finished Reports",
        operations_root / "03 Archive" / "processed-daily-reports",
        operations_root / "03 Archive" / "generated-workbooks",
        operations_root / "03 Archive" / "run-manifests",
    )
    for folder in required_folders:
        folder.mkdir(parents=True, exist_ok=True)
    lock_path = (
        operations_root
        / "03 Archive"
        / "run-manifests"
        / ".weekly-snapshot.lock"
    )
    lock_path.write_bytes(b"\0")
    anchor_dir = tmp_path / "integrity-anchors"
    anchor_dir.mkdir()
    (anchor_dir / "fixture-anchor.json").write_text(
        '{"schema_version":1}', encoding="utf-8"
    )

    import msvcrt

    with lock_path.open("r+b") as lock_handle:
        msvcrt.locking(lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
        try:
            concurrent = subprocess.run(
                [
                    POWERSHELL,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(repository / "Build-RecoveryBundle.ps1"),
                    "-DestinationDirectory",
                    str(tmp_path / "concurrent-recovery"),
                    "-IncludeOperationalData",
                    "-OperationsRoot",
                    str(operations_root),
                    "-IntegrityAnchorDirectory",
                    str(anchor_dir),
                ],
                cwd=repository,
                capture_output=True,
                text=True,
            )
        finally:
            lock_handle.seek(0)
            msvcrt.locking(lock_handle.fileno(), msvcrt.LK_UNLCK, 1)
    assert concurrent.returncode != 0
    assert "already running" in (concurrent.stdout + concurrent.stderr)

    operational_destination = tmp_path / "operational-recovery"
    operational = subprocess.run(
        [
            POWERSHELL,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(repository / "Build-RecoveryBundle.ps1"),
            "-DestinationDirectory",
            str(operational_destination),
            "-IncludeOperationalData",
            "-OperationsRoot",
            str(operations_root),
            "-IntegrityAnchorDirectory",
            str(anchor_dir),
        ],
        cwd=repository,
        capture_output=True,
        text=True,
    )
    assert operational.returncode == 0, operational.stdout + operational.stderr
    operational_bundles = list(operational_destination.glob("*.zip"))
    assert len(operational_bundles) == 1
    with zipfile.ZipFile(operational_bundles[0]) as archive:
        metadata = archive.read("release-metadata.json").decode("utf-8-sig")
        assert "operational_source_fingerprint_sha256" in metadata
        assert "restricted-operational-data.zip" in archive.namelist()
