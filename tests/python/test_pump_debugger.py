"""The pump keeps out of the way of VBA's debugger and editor.

A WM_TIMER that reaches ReDim's TIMERPROC while VBA is in break mode, or
after an edit to the module that holds it, takes Excel down. In 1.1.1 a Stop
in a click handler or in a job's step ended Excel about nine seconds into
break mode, and a comment typed into ReDimHost ended it in two. The
timer is now held while code ReDim runs, and the pump pauses while the VBA
editor or a VBA dialog is Excel's active window.

Break mode, and an editor brought forward, end an Excel session, so each
test runs in an Excel of its own. The VBA editor shows on screen while they
run, so they hold the lock compile checks take for it.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import sys
import time
from pathlib import Path

TOOLS = Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from build_workbooks import build_test_workbook  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402
from pyvbaharness.lock import COMPILE_MUTEX_NAME, SessionLock  # noqa: E402
from pyvbaharness.process_control import kill_pid_tree, wait_for_exit  # noqa: E402

# Before the fix, break mode with the timer armed ended Excel within 0.5 to
# 9 seconds, and typing into ReDimHost within 2.
BREAK_WATCH_S = 12.0
EDIT_WATCH_S = 6.0

WM_CHAR = 0x0102
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_RETURN = 0x0D
EDITOR_CLASS = "wndclass_desked_gsk"
CODE_WINDOW_CLASS = "VbaWindow"

user32 = ctypes.WinDLL("user32", use_last_error=True)
_ENUM_PROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
user32.EnumWindows.argtypes = [_ENUM_PROC, wt.LPARAM]
user32.EnumChildWindows.argtypes = [wt.HWND, _ENUM_PROC, wt.LPARAM]
user32.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetClassNameW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
user32.IsWindowVisible.argtypes = [wt.HWND]
user32.PostMessageW.argtypes = [wt.HWND, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
user32.PostMessageW.restype = wt.BOOL


def facts(session, step: str) -> dict[str, str]:
    value = session.run_raw(f"TestReDimAsync.{step}", timeout=60)
    return dict(pair.split("=", 1) for pair in str(value).split("|"))


def editor_shown() -> SessionLock:
    return SessionLock(timeout_s=600.0, name=COMPILE_MUTEX_NAME, purpose="compile")


def window_text(hwnd) -> str:
    buffer = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buffer, 512)
    return buffer.value


def window_class(hwnd) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def in_break_mode(pid: int) -> bool:
    """Whether a VBA editor window of the process says "[break]"."""
    found = []

    def visit(hwnd, _):
        owner = wt.DWORD(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
        if (owner.value == pid and user32.IsWindowVisible(hwnd)
                and window_class(hwnd) == EDITOR_CLASS
                and "[break]" in window_text(hwnd).lower()):
            found.append(hwnd)
        return True

    user32.EnumWindows(_ENUM_PROC(visit), 0)
    return bool(found)


def code_window(editor: int, module: str):
    found = []

    def visit(hwnd, _):
        if (window_class(hwnd) == CODE_WINDOW_CLASS
                and f"{module} (Code)" in window_text(hwnd)):
            found.append(hwnd)
        return True

    user32.EnumChildWindows(editor, _ENUM_PROC(visit), 0)
    assert found, f"the editor shows {module}'s code"
    return found[0]


def type_line(window, text: str) -> None:
    """Posts each character and then Enter, as typing in the editor sends them."""
    for character in text:
        user32.PostMessageW(window, WM_CHAR, ord(character), 1)
    user32.PostMessageW(window, WM_KEYDOWN, VK_RETURN, 0x001C0001)
    user32.PostMessageW(window, WM_KEYUP, VK_RETURN, 0xC01C0001)


def stops_without_ending(session, trigger: str) -> tuple[bool, bool]:
    """Runs the trigger, then watches Excel through break mode: whether it
    ended, and whether the debugger had stopped. Excel in break mode cannot
    quit, so it is ended here and the session closes at once."""
    session.run_raw(f"TestReDimAsync.{trigger}", timeout=60)
    exited = wait_for_exit(session.excel_pid, BREAK_WATCH_S)
    stopped = not exited and in_break_mode(session.excel_pid)
    kill_pid_tree(session.excel_pid)
    return exited, stopped


def test_stop_in_a_click_handler_leaves_excel_running():
    with ExcelSession() as session, editor_shown():
        session.open_workbook(str(build_test_workbook()))
        assert facts(session, "DebuggerSetup")["armed"] == "True", "the real timer runs"
        exited, stopped = stops_without_ending(session, "DebuggerClickLater")
    assert not exited, "Excel ended while the click handler sat in break mode"
    assert stopped, "the click handler stopped in the debugger"


def test_stop_in_a_job_step_leaves_excel_running():
    with ExcelSession() as session, editor_shown():
        session.open_workbook(str(build_test_workbook()))
        assert facts(session, "DebuggerSetup")["armed"] == "True", "the real timer runs"
        exited, stopped = stops_without_ending(session, "DebuggerStopInStep")
    assert not exited, "Excel ended while the job's step sat in break mode"
    assert stopped, "the step stopped in the debugger"


def test_typing_in_the_host_module_leaves_excel_running():
    with ExcelSession() as session, editor_shown():
        session.open_workbook(str(build_test_workbook()))
        assert facts(session, "DebuggerSetup")["armed"] == "True", "the real timer runs"
        editor = int(session.run_raw("TestReDimAsync.DebuggerOpenHostCode", timeout=60))
        time.sleep(0.5)
        state = facts(session, "DebuggerState")
        type_line(code_window(editor, "ReDimHost"), "' typed")
        exited = wait_for_exit(session.excel_pid, EDIT_WATCH_S)
    assert not exited, "Excel ended after an edit to ReDimHost"
    assert (state["armed"], state["paused"]) == ("False", "True"), (
        "the pump pauses while the editor is Excel's active window"
    )


def test_pump_pauses_for_the_editor():
    """The real editor made Excel's active window pauses the pump. A
    background Excel cannot hand activation back to a workbook window by
    itself, so test_async's test_pump_pauses covers the way back."""
    with ExcelSession() as session, editor_shown():
        session.open_workbook(str(build_test_workbook()))
        assert facts(session, "DebuggerSetup")["armed"] == "True", "the real timer runs"
        session.run_raw("TestReDimAsync.DebuggerShowExcel", timeout=60)
        time.sleep(0.3)
        shown = facts(session, "DebuggerState")
        session.run_raw("TestReDimAsync.DebuggerFocusEditor", timeout=60)
        time.sleep(0.3)
        paused = facts(session, "DebuggerState")
        time.sleep(0.3)
        still = facts(session, "DebuggerState")
        session.run_raw("TestReDimAsync.DebuggerSelectCell", timeout=60)
        selected = facts(session, "DebuggerState")
        teardown = facts(session, "DebuggerTeardown")
    assert (shown["armed"], shown["active"]) == ("True", "XLMAIN"), (
        "a workbook window active keeps the pump running"
    )
    assert (paused["armed"], paused["paused"], paused["active"]) == (
        "False", "True", EDITOR_CLASS), "the editor active pauses the pump"
    assert still["ticks"] == paused["ticks"], "no tick runs while it waits"
    assert (selected["armed"], selected["paused"]) == ("False", "True"), (
        "a selection while the editor stays active leaves the pump waiting"
    )
    assert teardown["armed"] == "False"
