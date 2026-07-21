from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture(autouse=True)
def isolate_integrity_anchor_state(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Keep trusted-head state deterministic and outside each test operator root."""

    monkeypatch.setenv(
        "RED_ONION_INTEGRITY_ANCHOR_DIR",
        str(tmp_path.parent / "trusted-integrity-anchors"),
    )
