"""Keys typed on a protected app surface, delivered as Excel gets them.

The test posts key messages to its own Excel's grid window, so they pass
through Excel's key handling and ReDim's OnKey capture as typing does. A
stray key with nothing focused must raise no protected-cell notice: the
harness reports such a modal and ends the session. A key typed while an
unlocked cell is active reaches the cell, and an app's hot key on a typing
key still runs.

Posting keys needs Excel idle between steps, and a failure ends the
session with a modal, so the file runs in an Excel session of its own.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import sys
import time
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from build_workbooks import build_test_workbook  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402
from pyvbaharness.results import HarnessError, SessionDead  # noqa: E402

WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_BACK = 0x08
VK_RETURN = 0x0D
VK_DELETE = 0x2E

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.PostMessageW.argtypes = [wt.HWND, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
user32.PostMessageW.restype = wt.BOOL
user32.MapVirtualKeyW.argtypes = [ctypes.c_uint, ctypes.c_uint]
user32.MapVirtualKeyW.restype = ctypes.c_uint


def press(grid: int, *keys: int) -> None:
    """Posts each key's down and up to the grid, then lets Excel take them."""
    for key in keys:
        scan = user32.MapVirtualKeyW(key, 0)
        extended = 1 << 24 if key == VK_DELETE else 0
        bits = 1 | (scan << 16) | extended
        user32.PostMessageW(grid, WM_KEYDOWN, key, bits)
        user32.PostMessageW(grid, WM_KEYUP, key, bits | (1 << 30) | (1 << 31))
        time.sleep(0.1)
    time.sleep(0.4)


def state(session, step: str, *args) -> dict[str, str]:
    try:
        value = session.run_raw(f"TestReDimCore.{step}", *args, timeout=60)
    except (SessionDead, HarnessError) as err:
        pytest.fail(f"Excel refused the step with a modal up, the protected-cell notice "
                    f"likely: {err}")
    return dict(pair.split("=", 1) for pair in str(value).split("|"))


def test_keys_on_a_protected_surface():
    workbook = build_test_workbook()
    with ExcelSession() as session:
        session.open_workbook(str(workbook))
        grid = int(session.run_raw("TestReDimCore.SurfaceKeysSetup", timeout=60))
        try:
            type_on_the_surface(session, grid)
        except BaseException:
            stop_pump_after_failure(session)
            raise
        # The pump stops before the session closes the workbook under it.
        assert session.run_raw("TestReDimCore.SurfaceKeysTeardown", timeout=60) == "False"


def stop_pump_after_failure(session) -> None:
    """Stops the pump once a step failed, when Excel still takes calls. A
    notice left up refuses the call; the session then ends Excel, and the
    timer with it, and the step's own failure is the one reported."""
    if session.is_dead:
        return
    try:
        session.run_raw("TestReDimCore.SurfaceKeysTeardown", timeout=60)
    except HarnessError:
        pass


def type_on_the_surface(session, grid: int) -> None:
    assert grid != 0, "the grid window is found"
    facts = state(session, "SurfaceKeysState")
    assert facts["pump"] == "True", "the pump runs, as in a live session"

    press(grid, ord("X"), VK_BACK, VK_DELETE, ord("5"))
    facts = state(session, "SurfaceKeysState")
    assert facts["held"] == "True", "the surface holds the typing keys"
    assert facts["A1"] == "", "stray keys end at the surface, with no notice"

    press(grid, ord("Q"))
    facts = state(session, "SurfaceKeysState")
    assert facts["hotKey"] == "1", "a hot key on a typing key still runs"

    state(session, "SurfaceKeysFocusField")
    press(grid, ord("H"), ord("I"), ord("Q"))
    facts = state(session, "SurfaceKeysState")
    assert facts["float"] == "hiq", "a focused field takes the keys, a hot key's too"
    assert facts["hotKey"] == "1", "the hot key waits while the field has its key"
    press(grid, VK_RETURN, ord("Z"))
    facts = state(session, "SurfaceKeysState")
    assert (facts["held"], facts["captured"]) == ("True", "False"), (
        "Enter hands the typing keys back to the surface"
    )
    assert facts["float"] == "hiq", "a key after Enter ends at the surface"
    press(grid, ord("Q"))
    facts = state(session, "SurfaceKeysState")
    assert facts["hotKey"] == "2", "the hot key has its key back after the field"

    facts = state(session, "SurfaceKeysUnlockedCell")
    assert facts["held"] == "True", "no event reported the unlocked cell"
    press(grid, ord("7"), VK_RETURN)
    facts = state(session, "SurfaceKeysState")
    assert facts["C7"] == "7", "a key typed into an unlocked cell reaches it"
    assert facts["held"] == "False", "and the surface lets the typing keys go"
    assert facts["faults"] == "0"
