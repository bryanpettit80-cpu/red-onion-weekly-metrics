"""Integrity primitives for the Red Onion weekly metrics workflow.

The reporting program owns workflow policy.  This module intentionally contains
only small, reusable building blocks for hashing, manifest persistence, manifest
chaining, inventory verification, and provenance collection.
"""

from __future__ import annotations

import hashlib
import importlib.metadata
import json
import os
import platform
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union


PathLike = Union[str, os.PathLike]
SHA256_PATTERN = re.compile(r"^[0-9a-fA-F]{64}$")
REQUIREMENT_NAME_PATTERN = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9._-]*)")
DEFAULT_HASH_CHUNK_SIZE = 1024 * 1024


class IntegrityError(RuntimeError):
    """Base exception for integrity failures."""


class PathEscapeError(IntegrityError):
    """Raised when a managed path resolves outside its declared root."""


class ManifestFormatError(IntegrityError):
    """Raised when a manifest is missing or malformed."""


class ManifestChainError(IntegrityError):
    """Raised when a previous-manifest chain cannot be verified."""


@dataclass(frozen=True)
class FileFingerprint:
    """A stable manifest record for one file below an inventory root."""

    path: str
    size: int
    sha256: str

    def __post_init__(self) -> None:
        normalized_path = normalize_relative_path(self.path)
        if isinstance(self.size, bool) or not isinstance(self.size, int) or self.size < 0:
            raise ValueError("File fingerprint size must be a non-negative integer.")
        if not isinstance(self.sha256, str) or not SHA256_PATTERN.fullmatch(self.sha256):
            raise ValueError("File fingerprint sha256 must be exactly 64 hexadecimal characters.")
        object.__setattr__(self, "path", normalized_path)
        object.__setattr__(self, "sha256", self.sha256.lower())

    def to_dict(self) -> Dict[str, Any]:
        return {"path": self.path, "size": self.size, "sha256": self.sha256}

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "FileFingerprint":
        try:
            path = value["path"]
            size = value["size"]
            digest = value["sha256"]
        except KeyError as exc:
            raise ValueError(f"File fingerprint is missing required field {exc.args[0]!r}.") from exc
        if not isinstance(path, str):
            raise ValueError("File fingerprint path must be a string.")
        return cls(path=path, size=size, sha256=digest)


@dataclass(frozen=True)
class InventoryComparison:
    """Differences between an expected and actual raw-file inventory."""

    missing: Tuple[FileFingerprint, ...] = ()
    unexpected: Tuple[FileFingerprint, ...] = ()
    changed: Tuple[Tuple[FileFingerprint, FileFingerprint], ...] = ()

    @property
    def ok(self) -> bool:
        return not (self.missing or self.unexpected or self.changed)


class RawInventoryError(IntegrityError):
    """Raised when raw files no longer match their recorded inventory."""

    def __init__(self, comparison: InventoryComparison):
        self.comparison = comparison
        super().__init__(format_inventory_mismatch(comparison))


@dataclass(frozen=True)
class ManifestChainEntry:
    """One verified entry in a newest-to-oldest manifest chain."""

    path: Path
    sha256: str
    payload: Mapping[str, Any]


def canonical_json_bytes(value: Any) -> bytes:
    """Serialize JSON deterministically for hashing.

    NaN and infinity are rejected because they are not portable JSON values.
    """

    serialized = json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return serialized.encode("utf-8")


def canonical_json_sha256(value: Any) -> str:
    """Return the SHA-256 digest of a value's canonical JSON representation."""

    return hashlib.sha256(canonical_json_bytes(value)).hexdigest()


def _hash_file_and_size(path: Path, chunk_size: int) -> Tuple[str, int]:
    if chunk_size <= 0:
        raise ValueError("chunk_size must be greater than zero.")
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    before = path.stat()
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)
    after = path.stat()

    before_identity = (
        before.st_size,
        before.st_mtime_ns,
        getattr(before, "st_dev", None),
        getattr(before, "st_ino", None),
    )
    after_identity = (
        after.st_size,
        after.st_mtime_ns,
        getattr(after, "st_dev", None),
        getattr(after, "st_ino", None),
    )
    if before_identity != after_identity:
        raise IntegrityError(f"File changed while it was being hashed: {path}")
    return digest.hexdigest(), after.st_size


def sha256_file(path: PathLike, chunk_size: int = DEFAULT_HASH_CHUNK_SIZE) -> str:
    """Hash a file without loading it entirely into memory."""

    digest, _ = _hash_file_and_size(Path(path), chunk_size)
    return digest


def normalize_relative_path(value: str) -> str:
    """Return a platform-neutral safe relative path for a manifest.

    Windows drive paths, UNC paths, traversal, alternate-data-stream syntax,
    and ambiguous trailing dots/spaces are rejected even on non-Windows test
    hosts so a manifest has the same safety semantics everywhere.
    """

    if not isinstance(value, str):
        raise PathEscapeError("Manifest path must be a string.")
    raw = value.replace("\\", "/")
    if not raw or "\x00" in raw:
        raise PathEscapeError(f"Unsafe relative path: {value!r}")

    posix_path = PurePosixPath(raw)
    windows_path = PureWindowsPath(raw)
    if posix_path.is_absolute() or windows_path.is_absolute() or windows_path.drive:
        raise PathEscapeError(f"Unsafe relative path: {value!r}")

    parts = raw.split("/")
    if any(
        not part
        or part in {".", ".."}
        or ":" in part
        or part != part.rstrip(" .")
        for part in parts
    ):
        raise PathEscapeError(f"Unsafe relative path: {value!r}")
    return "/".join(parts)


def resolve_relative_path(root: PathLike, relative_path: str) -> Path:
    """Resolve a safe manifest-relative path and enforce its root boundary."""

    normalized = normalize_relative_path(relative_path)
    root_path = Path(root).resolve()
    candidate = (root_path / Path(*normalized.split("/"))).resolve()
    try:
        candidate.relative_to(root_path)
    except ValueError as exc:
        raise PathEscapeError(
            f"Path escapes managed root {root_path}: {relative_path!r}"
        ) from exc
    return candidate


def path_within_root(root: PathLike, path: PathLike) -> Path:
    """Resolve an absolute or root-relative path and enforce the root boundary."""

    root_path = Path(root).resolve()
    supplied = Path(path)
    if supplied.is_absolute():
        candidate = supplied.resolve()
    else:
        # Apply manifest-grade normalization to relative caller input as well.
        # This also catches Windows absolute/drive paths on a non-Windows test host.
        normalized = normalize_relative_path(os.fspath(path))
        candidate = (root_path / Path(*normalized.split("/"))).resolve()
    try:
        candidate.relative_to(root_path)
    except ValueError as exc:
        raise PathEscapeError(f"Path escapes managed root {root_path}: {path!s}") from exc
    return candidate


def relative_path_within_root(root: PathLike, path: PathLike) -> str:
    """Return a safe POSIX-style path relative to root."""

    root_path = Path(root).resolve()
    candidate = path_within_root(root_path, path)
    relative = candidate.relative_to(root_path).as_posix()
    return normalize_relative_path(relative)


def fingerprint_file(root: PathLike, path: PathLike) -> FileFingerprint:
    """Fingerprint one file and record its safe path relative to root."""

    root_path = Path(root).resolve()
    candidate = path_within_root(root_path, path)
    relative = relative_path_within_root(root_path, candidate)
    digest, size = _hash_file_and_size(candidate, DEFAULT_HASH_CHUNK_SIZE)
    return FileFingerprint(path=relative, size=size, sha256=digest)


def _casefold_key(path: str) -> str:
    return path.casefold()


def _fingerprint_map(
    values: Iterable[Union[FileFingerprint, Mapping[str, Any]]]
) -> Dict[str, FileFingerprint]:
    result: Dict[str, FileFingerprint] = {}
    for value in values:
        fingerprint = (
            value if isinstance(value, FileFingerprint) else FileFingerprint.from_mapping(value)
        )
        key = _casefold_key(fingerprint.path)
        if key in result:
            raise ValueError(f"Duplicate inventory path: {fingerprint.path}")
        result[key] = fingerprint
    return result


def build_raw_inventory(
    root: PathLike,
    paths: Optional[Iterable[PathLike]] = None,
) -> Tuple[FileFingerprint, ...]:
    """Build a sorted inventory for explicit files or every file below root."""

    root_path = Path(root).resolve()
    candidates: Iterable[PathLike]
    if paths is None:
        candidates = (path for path in root_path.rglob("*") if path.is_file())
    else:
        candidates = paths

    fingerprints = [fingerprint_file(root_path, path) for path in candidates]
    mapped = _fingerprint_map(fingerprints)
    return tuple(sorted(mapped.values(), key=lambda item: _casefold_key(item.path)))


def compare_raw_inventory(
    expected: Iterable[Union[FileFingerprint, Mapping[str, Any]]],
    actual: Iterable[Union[FileFingerprint, Mapping[str, Any]]],
) -> InventoryComparison:
    """Compare two inventories without touching the filesystem."""

    expected_map = _fingerprint_map(expected)
    actual_map = _fingerprint_map(actual)
    missing: List[FileFingerprint] = []
    unexpected: List[FileFingerprint] = []
    changed: List[Tuple[FileFingerprint, FileFingerprint]] = []

    for key in sorted(expected_map):
        expected_item = expected_map[key]
        actual_item = actual_map.get(key)
        if actual_item is None:
            missing.append(expected_item)
        elif (
            expected_item.size != actual_item.size
            or expected_item.sha256 != actual_item.sha256
            or expected_item.path != actual_item.path
        ):
            changed.append((expected_item, actual_item))

    for key in sorted(actual_map):
        if key not in expected_map:
            unexpected.append(actual_map[key])

    return InventoryComparison(
        missing=tuple(missing),
        unexpected=tuple(unexpected),
        changed=tuple(changed),
    )


def format_inventory_mismatch(comparison: InventoryComparison) -> str:
    """Format stable expected-versus-actual diagnostics for operators and tests."""

    lines = ["Raw inventory verification failed:"]
    for item in comparison.missing:
        lines.append(
            f"  - {item.path}: expected(size={item.size}, sha256={item.sha256}); actual(missing)"
        )
    for item in comparison.unexpected:
        lines.append(
            f"  - {item.path}: expected(missing); actual(size={item.size}, sha256={item.sha256})"
        )
    for expected_item, actual_item in comparison.changed:
        lines.append(
            f"  - {expected_item.path}: "
            f"expected(size={expected_item.size}, sha256={expected_item.sha256}); "
            f"actual(size={actual_item.size}, sha256={actual_item.sha256})"
        )
    return "\n".join(lines)


def verify_raw_inventory(
    root: PathLike,
    expected: Iterable[Union[FileFingerprint, Mapping[str, Any]]],
    actual_paths: Optional[Iterable[PathLike]] = None,
) -> Tuple[FileFingerprint, ...]:
    """Hash the current raw files and require an exact inventory match."""

    actual = build_raw_inventory(root, actual_paths)
    comparison = compare_raw_inventory(expected, actual)
    if not comparison.ok:
        raise RawInventoryError(comparison)
    return actual


def read_json_manifest(path: PathLike, root: Optional[PathLike] = None) -> Dict[str, Any]:
    """Read and validate a JSON manifest object."""

    manifest_path = path_within_root(root, path) if root is not None else Path(path).resolve()

    def reject_nonstandard_constant(value: str) -> None:
        raise ValueError(f"non-standard numeric constant {value}")

    try:
        with manifest_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle, parse_constant=reject_nonstandard_constant)
    except FileNotFoundError as exc:
        raise ManifestFormatError(f"Manifest not found: {manifest_path}") from exc
    except json.JSONDecodeError as exc:
        raise ManifestFormatError(
            f"Invalid JSON manifest {manifest_path}: {exc.msg} "
            f"at line {exc.lineno} column {exc.colno}"
        ) from exc
    except (UnicodeError, ValueError) as exc:
        raise ManifestFormatError(f"Invalid JSON manifest {manifest_path}: {exc}") from exc
    except OSError as exc:
        raise ManifestFormatError(f"Could not read manifest {manifest_path}: {exc}") from exc
    if not isinstance(payload, dict):
        raise ManifestFormatError(f"Manifest must contain a JSON object: {manifest_path}")
    return payload


def write_json_manifest_atomic(
    path: PathLike,
    payload: Mapping[str, Any],
    root: Optional[PathLike] = None,
) -> Path:
    """Write a JSON manifest through a same-directory temporary file and replace."""

    if not isinstance(payload, Mapping):
        raise TypeError("Manifest payload must be a mapping.")
    manifest_path = path_within_root(root, path) if root is not None else Path(path).resolve()
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    serialized = json.dumps(
        payload,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        indent=2,
    ) + "\n"

    temp_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=f".{manifest_path.name}.",
            suffix=".tmp",
            dir=str(manifest_path.parent),
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            handle.write(serialized)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(str(temp_path), str(manifest_path))
        temp_path = None
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass
    return manifest_path


def link_previous_manifest(
    payload: Mapping[str, Any],
    manifest_root: PathLike,
    previous_manifest_path: PathLike,
    expected_previous_sha256: Optional[str] = None,
) -> Dict[str, Any]:
    """Return a copied payload linked to the canonical hash of its predecessor."""

    if "previous_manifest" in payload:
        raise ManifestFormatError("Manifest payload already contains previous_manifest.")
    root_path = Path(manifest_root).resolve()
    previous_path = path_within_root(root_path, previous_manifest_path)
    previous_relative = relative_path_within_root(root_path, previous_path)
    previous_payload = read_json_manifest(previous_path, root=root_path)
    previous_sha256 = canonical_json_sha256(previous_payload)
    if expected_previous_sha256 is not None:
        expected = _validated_sha256(
            expected_previous_sha256, "Expected previous manifest sha256"
        )
        if previous_sha256 != expected:
            raise ManifestChainError(
                "The previous integrity manifest changed after preflight; refusing to "
                "extend a different history."
            )
    linked = dict(payload)
    linked["previous_manifest"] = {
        "path": previous_relative,
        "sha256": previous_sha256,
    }
    return linked


def write_chained_manifest_atomic(
    path: PathLike,
    payload: Mapping[str, Any],
    manifest_root: PathLike,
    previous_manifest_path: Optional[PathLike] = None,
    expected_previous_sha256: Optional[str] = None,
) -> Path:
    """Optionally link a manifest to its predecessor, then write it atomically."""

    linked_payload = (
        link_previous_manifest(
            payload,
            manifest_root,
            previous_manifest_path,
            expected_previous_sha256,
        )
        if previous_manifest_path is not None
        else dict(payload)
    )
    return write_json_manifest_atomic(path, linked_payload, root=manifest_root)


def _validated_sha256(value: Any, context: str) -> str:
    if not isinstance(value, str) or not SHA256_PATTERN.fullmatch(value):
        raise ManifestChainError(f"{context} must be a 64-character SHA-256 digest.")
    return value.lower()


def verify_manifest_chain(
    latest_manifest_path: PathLike,
    manifest_root: Optional[PathLike] = None,
) -> Tuple[ManifestChainEntry, ...]:
    """Verify every previous-manifest link and return newest-to-oldest entries."""

    if manifest_root is None:
        latest_supplied = Path(latest_manifest_path).resolve()
        root_path = latest_supplied.parent
        current_path = latest_supplied
    else:
        root_path = Path(manifest_root).resolve()
        current_path = path_within_root(root_path, latest_manifest_path)

    entries: List[ManifestChainEntry] = []
    seen: set = set()
    while True:
        current_relative = relative_path_within_root(root_path, current_path)
        current_key = _casefold_key(current_relative)
        if current_key in seen:
            raise ManifestChainError(
                f"Manifest chain contains a cycle at {current_relative}."
            )
        seen.add(current_key)

        payload = read_json_manifest(current_path, root=root_path)
        current_hash = canonical_json_sha256(payload)
        entries.append(
            ManifestChainEntry(path=current_path, sha256=current_hash, payload=payload)
        )

        previous = payload.get("previous_manifest")
        if previous is None:
            break
        if not isinstance(previous, dict):
            raise ManifestChainError(
                f"Manifest {current_relative} has an invalid previous_manifest reference."
            )
        previous_relative = previous.get("path")
        if not isinstance(previous_relative, str):
            raise ManifestChainError(
                f"Manifest {current_relative} previous_manifest.path must be a string."
            )
        expected_hash = _validated_sha256(
            previous.get("sha256"),
            f"Manifest {current_relative} previous_manifest.sha256",
        )
        try:
            previous_path = resolve_relative_path(root_path, previous_relative)
        except PathEscapeError as exc:
            raise ManifestChainError(
                f"Manifest {current_relative} has an unsafe previous path {previous_relative!r}."
            ) from exc
        previous_normalized = relative_path_within_root(root_path, previous_path)
        if _casefold_key(previous_normalized) in seen:
            raise ManifestChainError(
                f"Manifest chain contains a cycle at {previous_normalized}."
            )
        if not previous_path.is_file():
            raise ManifestChainError(
                f"Manifest chain is missing previous manifest {previous_normalized} "
                f"referenced by {current_relative}."
            )
        previous_payload = read_json_manifest(previous_path, root=root_path)
        actual_hash = canonical_json_sha256(previous_payload)
        if actual_hash != expected_hash:
            raise ManifestChainError(
                f"Manifest chain hash mismatch for {previous_normalized}: "
                f"expected sha256={expected_hash}; actual sha256={actual_hash}."
            )
        current_path = previous_path

    return tuple(entries)


def _git_command(repo_root: Path, arguments: Sequence[str]) -> Tuple[bool, str]:
    try:
        completed = subprocess.run(
            ["git", "-C", str(repo_root), *arguments],
            check=False,
            capture_output=True,
            text=True,
            timeout=5,
        )
    except FileNotFoundError:
        return False, "git executable not found"
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, str(exc)
    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip() or (
            f"git exited with code {completed.returncode}"
        )
        return False, detail.splitlines()[0]
    return True, completed.stdout.strip()


def collect_git_provenance(repo_root: PathLike) -> Dict[str, Any]:
    """Collect Git revision/branch/cleanliness without mutating the repository."""

    repo_path = Path(repo_root).resolve()
    ok, head = _git_command(repo_path, ["rev-parse", "HEAD"])
    if not ok:
        return {"available": False, "error": head}

    branch_ok, branch = _git_command(repo_path, ["rev-parse", "--abbrev-ref", "HEAD"])
    status_ok, status = _git_command(
        repo_path, ["status", "--porcelain=v1", "--untracked-files=no"]
    )
    result: Dict[str, Any] = {
        "available": True,
        "commit": head,
        "branch": branch if branch_ok else None,
        "tracked_files_clean": not bool(status) if status_ok else None,
    }
    if not branch_ok:
        result["branch_error"] = branch
    if not status_ok:
        result["status_error"] = status
    return result


def dependency_names_from_requirements(path: PathLike) -> Tuple[str, ...]:
    """Extract simple distribution names from a requirements file."""

    names: Dict[str, str] = {}
    with Path(path).open("r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.split("#", 1)[0].strip()
            if not line or line.startswith(("-", "http://", "https://", "git+")):
                continue
            match = REQUIREMENT_NAME_PATTERN.match(line)
            if match:
                name = match.group(1)
                names.setdefault(name.casefold(), name)
    return tuple(names[key] for key in sorted(names))


def collect_dependency_versions(distributions: Iterable[str]) -> Dict[str, Optional[str]]:
    """Return installed versions, using null for distributions not installed."""

    names: Dict[str, str] = {}
    for value in distributions:
        name = str(value).strip()
        if not name:
            continue
        names.setdefault(name.casefold(), name)
    versions: Dict[str, Optional[str]] = {}
    for key in sorted(names):
        name = names[key]
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = None
    return versions


def _file_provenance(repo_root: Path, path: PathLike) -> Dict[str, Any]:
    candidate = path_within_root(repo_root, path)
    relative = relative_path_within_root(repo_root, candidate)
    if not candidate.is_file():
        return {"path": relative, "exists": False}
    digest, size = _hash_file_and_size(candidate, DEFAULT_HASH_CHUNK_SIZE)
    return {"path": relative, "exists": True, "size": size, "sha256": digest}


def collect_provenance(
    repo_root: PathLike,
    config_path: Optional[PathLike] = None,
    requirements_path: Optional[PathLike] = None,
    dependency_names: Optional[Iterable[str]] = None,
) -> Dict[str, Any]:
    """Collect reproducibility metadata without requiring Git or every package."""

    repo_path = Path(repo_root).resolve()
    result: Dict[str, Any] = {
        "git": collect_git_provenance(repo_path),
        "python": {
            "implementation": platform.python_implementation(),
            "version": platform.python_version(),
            "executable": sys.executable,
            "platform": platform.platform(),
        },
    }

    if config_path is not None:
        result["config"] = _file_provenance(repo_path, config_path)

    declared_dependencies: Tuple[str, ...] = ()
    if requirements_path is not None:
        requirements = path_within_root(repo_path, requirements_path)
        result["requirements"] = _file_provenance(repo_path, requirements)
        if requirements.is_file():
            declared_dependencies = dependency_names_from_requirements(requirements)

    selected_dependencies = (
        tuple(dependency_names) if dependency_names is not None else declared_dependencies
    )
    result["dependencies"] = collect_dependency_versions(selected_dependencies)
    return result


__all__ = [
    "FileFingerprint",
    "IntegrityError",
    "InventoryComparison",
    "ManifestChainEntry",
    "ManifestChainError",
    "ManifestFormatError",
    "PathEscapeError",
    "RawInventoryError",
    "build_raw_inventory",
    "canonical_json_bytes",
    "canonical_json_sha256",
    "collect_dependency_versions",
    "collect_git_provenance",
    "collect_provenance",
    "compare_raw_inventory",
    "dependency_names_from_requirements",
    "fingerprint_file",
    "format_inventory_mismatch",
    "link_previous_manifest",
    "normalize_relative_path",
    "path_within_root",
    "read_json_manifest",
    "relative_path_within_root",
    "resolve_relative_path",
    "sha256_file",
    "verify_manifest_chain",
    "verify_raw_inventory",
    "write_chained_manifest_atomic",
    "write_json_manifest_atomic",
]
