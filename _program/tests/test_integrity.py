import hashlib
import json
from pathlib import Path

import pytest

import red_onion_integrity as integrity


def test_canonical_json_hash_is_order_independent_and_rejects_nan() -> None:
    first = {"z": [3, 2, 1], "a": {"name": "Caf\u00e9", "enabled": True}}
    second = {"a": {"enabled": True, "name": "Caf\u00e9"}, "z": [3, 2, 1]}

    assert integrity.canonical_json_bytes(first) == integrity.canonical_json_bytes(second)
    assert integrity.canonical_json_sha256(first) == integrity.canonical_json_sha256(second)
    with pytest.raises(ValueError, match="Out of range float values"):
        integrity.canonical_json_bytes({"bad": float("nan")})


def test_sha256_file_streams_expected_digest(tmp_path: Path) -> None:
    source = tmp_path / "daily.xlsx"
    content = (b"red-onion\x00" * 200_000) + b"end"
    source.write_bytes(content)

    assert integrity.sha256_file(source, chunk_size=31) == hashlib.sha256(content).hexdigest()
    with pytest.raises(ValueError, match="chunk_size must be greater than zero"):
        integrity.sha256_file(source, chunk_size=0)


@pytest.mark.parametrize(
    "unsafe",
    [
        "../outside.json",
        "folder/../outside.json",
        "/absolute.json",
        r"C:\outside.json",
        r"\\server\share\outside.json",
        "folder//file.json",
        "file.json:stream",
        "trailing. /file.json",
    ],
)
def test_manifest_paths_reject_escape_and_windows_ambiguities(
    tmp_path: Path, unsafe: str
) -> None:
    with pytest.raises(integrity.PathEscapeError, match="Unsafe relative path"):
        integrity.resolve_relative_path(tmp_path, unsafe)


def test_resolved_symlink_escape_is_rejected_when_supported(tmp_path: Path) -> None:
    root = tmp_path / "root"
    outside = tmp_path / "outside"
    root.mkdir()
    outside.mkdir()
    link = root / "linked"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("Directory symlinks are unavailable on this host")

    with pytest.raises(integrity.PathEscapeError, match="escapes managed root"):
        integrity.resolve_relative_path(root, "linked/manifest.json")


def test_general_root_boundary_rejects_relative_traversal_and_foreign_drive(
    tmp_path: Path,
) -> None:
    with pytest.raises(integrity.PathEscapeError, match="Unsafe relative path"):
        integrity.path_within_root(tmp_path, "folder/../file.json")
    with pytest.raises(
        integrity.PathEscapeError, match="Unsafe relative path|escapes managed root"
    ):
        integrity.path_within_root(tmp_path, r"Z:\foreign\file.json")


def test_atomic_manifest_round_trip_is_deterministic_and_leaves_no_temp(
    tmp_path: Path,
) -> None:
    path = tmp_path / "manifests" / "run.json"
    payload = {"version": 1, "raw": [{"path": "week/day.xlsx", "size": 3}]}

    written = integrity.write_json_manifest_atomic(path, payload, root=tmp_path)

    assert written == path.resolve()
    assert integrity.read_json_manifest(path, root=tmp_path) == payload
    assert path.read_text(encoding="utf-8") == json.dumps(
        payload, ensure_ascii=False, allow_nan=False, sort_keys=True, indent=2
    ) + "\n"
    assert list(path.parent.glob(f".{path.name}.*.tmp")) == []


def test_atomic_manifest_failure_preserves_existing_file_and_cleans_temp(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "run.json"
    path.write_text('{"state":"old"}\n', encoding="utf-8")

    def fail_replace(source: str, destination: str) -> None:
        raise OSError("simulated replace failure")

    monkeypatch.setattr(integrity.os, "replace", fail_replace)
    with pytest.raises(OSError, match="simulated replace failure"):
        integrity.write_json_manifest_atomic(path, {"state": "new"})

    assert path.read_text(encoding="utf-8") == '{"state":"old"}\n'
    assert list(tmp_path.glob(f".{path.name}.*.tmp")) == []


def test_manifest_read_reports_missing_invalid_and_non_object(tmp_path: Path) -> None:
    missing = tmp_path / "missing.json"
    with pytest.raises(integrity.ManifestFormatError, match="Manifest not found"):
        integrity.read_json_manifest(missing)

    invalid = tmp_path / "invalid.json"
    invalid.write_text("{not json", encoding="utf-8")
    with pytest.raises(integrity.ManifestFormatError, match=r"Invalid JSON manifest .* line 1 column 2"):
        integrity.read_json_manifest(invalid)

    array = tmp_path / "array.json"
    array.write_text("[]", encoding="utf-8")
    with pytest.raises(integrity.ManifestFormatError, match="must contain a JSON object"):
        integrity.read_json_manifest(array)

    nonstandard = tmp_path / "nonstandard.json"
    nonstandard.write_text('{"value": NaN}', encoding="utf-8")
    with pytest.raises(integrity.ManifestFormatError, match="non-standard numeric constant NaN"):
        integrity.read_json_manifest(nonstandard)


def test_raw_inventory_verification_reports_exact_expected_and_actual(
    tmp_path: Path,
) -> None:
    raw = tmp_path / "raw"
    raw.mkdir()
    changed = raw / "changed.xlsx"
    unexpected = raw / "unexpected.xls"
    changed.write_bytes(b"actual")
    unexpected.write_bytes(b"extra")

    expected = [
        integrity.FileFingerprint(
            path="changed.xlsx",
            size=len(b"expected"),
            sha256=hashlib.sha256(b"expected").hexdigest(),
        ),
        integrity.FileFingerprint(
            path="missing.xlsx",
            size=7,
            sha256=hashlib.sha256(b"missing").hexdigest(),
        ),
    ]
    actual_changed_hash = hashlib.sha256(b"actual").hexdigest()
    unexpected_hash = hashlib.sha256(b"extra").hexdigest()

    with pytest.raises(integrity.RawInventoryError) as exc_info:
        integrity.verify_raw_inventory(raw, expected)

    assert str(exc_info.value) == "\n".join(
        [
            "Raw inventory verification failed:",
            (
                "  - missing.xlsx: "
                f"expected(size=7, sha256={hashlib.sha256(b'missing').hexdigest()}); "
                "actual(missing)"
            ),
            (
                "  - unexpected.xls: expected(missing); "
                f"actual(size=5, sha256={unexpected_hash})"
            ),
            (
                "  - changed.xlsx: "
                f"expected(size=8, sha256={hashlib.sha256(b'expected').hexdigest()}); "
                f"actual(size=6, sha256={actual_changed_hash})"
            ),
        ]
    )


def test_raw_inventory_accepts_mapping_records_and_explicit_paths(tmp_path: Path) -> None:
    raw = tmp_path / "raw"
    raw.mkdir()
    included = raw / "included.xlsx"
    ignored = raw / "notes.txt"
    included.write_bytes(b"source")
    ignored.write_text("not part of the raw inventory", encoding="utf-8")
    expected = [integrity.fingerprint_file(raw, included).to_dict()]

    verified = integrity.verify_raw_inventory(raw, expected, actual_paths=[included])

    assert [item.path for item in verified] == ["included.xlsx"]


def test_two_manifest_chain_verifies_and_detects_predecessor_tampering(
    tmp_path: Path,
) -> None:
    root = tmp_path / "manifests"
    first = root / "001.json"
    second = root / "002.json"
    integrity.write_json_manifest_atomic(first, {"run": 1}, root=root)
    integrity.write_chained_manifest_atomic(
        second,
        {"run": 2},
        manifest_root=root,
        previous_manifest_path=first,
    )

    entries = integrity.verify_manifest_chain(second, manifest_root=root)

    assert [entry.path.name for entry in entries] == ["002.json", "001.json"]
    first.write_text('{"run": 999}\n', encoding="utf-8")
    with pytest.raises(
        integrity.ManifestChainError,
        match=r"hash mismatch for 001\.json: expected sha256=.*; actual sha256=.*",
    ):
        integrity.verify_manifest_chain(second, manifest_root=root)


def test_manifest_chain_reports_missing_cycle_and_path_escape(tmp_path: Path) -> None:
    root = tmp_path / "manifests"
    root.mkdir()
    digest = "0" * 64

    missing = root / "missing-link.json"
    integrity.write_json_manifest_atomic(
        missing,
        {"previous_manifest": {"path": "gone.json", "sha256": digest}},
        root=root,
    )
    with pytest.raises(
        integrity.ManifestChainError,
        match="missing previous manifest gone.json referenced by missing-link.json",
    ):
        integrity.verify_manifest_chain(missing, manifest_root=root)

    cycle = root / "cycle.json"
    integrity.write_json_manifest_atomic(
        cycle,
        {"previous_manifest": {"path": "cycle.json", "sha256": digest}},
        root=root,
    )
    with pytest.raises(integrity.ManifestChainError, match="contains a cycle at cycle.json"):
        integrity.verify_manifest_chain(cycle, manifest_root=root)

    escape = root / "escape.json"
    integrity.write_json_manifest_atomic(
        escape,
        {"previous_manifest": {"path": "../outside.json", "sha256": digest}},
        root=root,
    )
    with pytest.raises(
        integrity.ManifestChainError,
        match=r"unsafe previous path '\.\./outside\.json'",
    ):
        integrity.verify_manifest_chain(escape, manifest_root=root)


def test_provenance_hashes_managed_files_and_dependency_versions(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    program = repo / "_program"
    program.mkdir(parents=True)
    config = program / "config.json"
    requirements = program / "requirements.txt"
    config.write_text('{"locations": {}}\n', encoding="utf-8")
    requirements.write_text("pytest>=8\ndefinitely-not-installed-red-onion-package==1\n", encoding="utf-8")

    provenance = integrity.collect_provenance(
        repo,
        config_path=config,
        requirements_path=requirements,
    )

    assert provenance["config"] == {
        "path": "_program/config.json",
        "exists": True,
        "size": config.stat().st_size,
        "sha256": integrity.sha256_file(config),
    }
    assert provenance["requirements"]["sha256"] == integrity.sha256_file(requirements)
    assert provenance["dependencies"]["pytest"] is not None
    assert provenance["dependencies"]["definitely-not-installed-red-onion-package"] is None
    assert isinstance(provenance["git"]["available"], bool)
    assert provenance["python"]["version"]


def test_provenance_rejects_config_outside_repository(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    outside = tmp_path / "outside.json"
    outside.write_text("{}", encoding="utf-8")

    with pytest.raises(integrity.PathEscapeError, match="escapes managed root"):
        integrity.collect_provenance(repo, config_path=outside)
