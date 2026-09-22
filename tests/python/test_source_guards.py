"""Pure-source structural guards: invariants the VBA runtime must hold by
construction, checked by reading the source without launching Excel.

These lock in defect-class fixes at the shape level, not just the repaired
line, the same way ROneCOne guards its own repaired classes.
"""

from __future__ import annotations

import re
from pathlib import Path

from vba_sources import RELEASE_SOURCES, release_header

REDIMUI = Path(__file__).resolve().parents[2] / "src" / "ReDimUI.cls"

# A bare IsArray call, not IsArrayValue: "IsArray" followed by optional
# whitespace and an open paren. "IsArrayValue(" has "Value" after IsArray,
# so it never matches, and prose mentions of IsArray lack the paren.
ISARRAY_CALL = re.compile(r"\bIsArray\s*\(")
GUARD_FUNCTION = re.compile(
    r"Private Function IsArrayValue.*?End Function", re.DOTALL
)


def test_isarray_appears_once_inside_the_guard():
    text = REDIMUI.read_text(encoding="utf-8")
    calls = ISARRAY_CALL.findall(text)
    assert len(calls) == 1, (
        f"IsArray must appear exactly once in the runtime, found {len(calls)}. "
        "An object reaching a bare IsArray dereferences its default member and "
        "raises; route every array test through the IsArrayValue guard."
    )
    guard = GUARD_FUNCTION.search(text)
    assert guard is not None, "the IsArrayValue guard function must exist"
    assert len(ISARRAY_CALL.findall(guard.group(0))) == 1, (
        "the one IsArray call must live inside the IsArrayValue guard, "
        "which short-circuits on IsObject before testing"
    )


def test_release_sources_carry_license_and_version():
    """Every source a release ships opens with the version and release date,
    the repository, and the MIT license, built from REDIM_VERSION, the
    CHANGELOG, and LICENSE; tools/stamp_release.py writes it."""
    header = release_header()
    for path in RELEASE_SOURCES:
        text = path.read_text(encoding="utf-8").replace("\r\n", "\n")
        preamble = text.split("\nOption Explicit", 1)[0]
        assert header in preamble, (
            f"{path.name} lacks the current release header; run "
            "python tools/stamp_release.py:\n" + header
        )
