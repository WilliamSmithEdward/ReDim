"""Smoke tests: each demo workbook builds, mounts, and survives interaction.

Each demo runs in its own Excel session because demos own module-level game
or feed state and hook application keys.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[2] / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from build_workbooks import build_demo_workbooks  # noqa: E402
from vba_sources import ROOT  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402


@pytest.fixture(scope="module")
def demo_paths():
    return {path.name: path for path in build_demo_workbooks()}


def run(excel, macro, *args, timeout=90):
    result = excel.run_macro(macro, *args, timeout=timeout)
    assert result.outcome == "passed", (
        f"{macro} outcome={result.outcome} error={result.error}"
    )
    return result.value


SMOKE_MISSION = """
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function SmokeMission() As String
    Dim app As ReDimUI
    Dim ticks As Long
    Dim transcript As String

    ReDimUI.AutoPump False
    BuildMissionControl
    Set app = MissionApp()
    transcript = "components=" & app.ComponentCount

    HandleStartFeed1
    transcript = transcript & "|feed1Running=" & CStr(app.State("feed1Running"))
    transcript = transcript & "|launchDisabled=" & _
        CStr(Not app.Button("launch").IsEnabled)
    ' Feed 1 is paced at 80ms per step, so give each tick a pace window.
    For ticks = 1 To 6
        Sleep 100
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|progressMoved=" & _
        CStr(app.State("feed1Pct") > 0)
    transcript = transcript & "|rowsMoved=" & CStr(app.State("rowsLoaded") > 0)

    HandleCancelFeed1
    ReDimUI.PumpOnce
    transcript = transcript & "|canceledStatus=" & app.State("feed1Status")
    transcript = transcript & "|startReEnabled=" & CStr(app.State("feed1Idle"))
    transcript = transcript & "|launchReEnabled=" & _
        CStr(app.Button("launch").IsEnabled)

    ' Toasts sit on a rail just outside the content's right edge.
    Dim toastShape As Shape
    Dim cardShape As Shape
    Set cardShape = app.Sheet.Shapes("rdm_mission_card1")
    Set toastShape = app.Sheet.Shapes("rdm_mission_toast_1")
    transcript = transcript & "|toastOnRail=" & _
        CStr(Abs(toastShape.Left - _
            (cardShape.Left + cardShape.Width + 12)) < 1)
    transcript = transcript & "|toastInkOnSurface=" & _
        CStr(toastShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            app.Theme.OnSurfaceColor)
    transcript = transcript & "|surfaceProtected=" & _
        CStr(app.Sheet.ProtectContents)
    ReDimUI.AutoPump True
    SmokeMission = transcript
End Function
"""

SMOKE_NAVIGATOR = """
Public Function SmokeNavigator() As String
    Dim transcript As String

    ReDimUI.AutoPump False
    BuildNavigator
    ReDimUI.Navigate "navhome"
    transcript = "active=" & ReDimUI.ActiveWindowId
    transcript = transcript & "|homeVisible=" & _
        CStr(ThisWorkbook.Worksheets("NavHome").Visible = xlSheetVisible)
    transcript = transcript & "|settingsHidden=" & _
        CStr(ThisWorkbook.Worksheets("NavSettings").Visible = xlSheetVeryHidden)
    transcript = transcript & "|homeShownOnce=" & _
        CStr(ReDimUI.App("navhome").State("homeShown") = 1)
    transcript = transcript & "|homeProtected=" & _
        CStr(ThisWorkbook.Worksheets("NavHome").ProtectContents)

    transcript = transcript & "|homeTabActive=" & _
        CStr(ThisWorkbook.Worksheets("NavHome").Shapes( _
            "rdm_navhome_nvb_navhome").Fill.ForeColor.RGB = _
            ReDimUI.App("navhome").Theme.PrimaryColor)
    transcript = transcript & "|tabTitle=" & _
        ThisWorkbook.Worksheets("NavHome").Shapes( _
            "rdm_navhome_nvb_navsettings").TextFrame2.TextRange.Text

    ReDimUI.DispatchShape "rdm_navhome_nvb_navsettings"
    transcript = transcript & "|activeSettings=" & _
        CStr(ReDimUI.ActiveWindowId = "navsettings")
    transcript = transcript & "|homeNowHidden=" & _
        CStr(ThisWorkbook.Worksheets("NavHome").Visible = xlSheetVeryHidden)
    transcript = transcript & "|settingsTabActive=" & _
        CStr(ThisWorkbook.Worksheets("NavSettings").Shapes( _
            "rdm_navsettings_nvb_navsettings").Fill.ForeColor.RGB = _
            ReDimUI.App("navsettings").Theme.PrimaryColor)
    transcript = transcript & "|persistedDefault=" & _
        CStr(ReDimUI.App("navsettings").State("alertsOn"))

    HandleBack
    transcript = transcript & "|backHome=" & _
        CStr(ReDimUI.ActiveWindowId = "navhome")
    transcript = transcript & "|homeShownTwice=" & _
        CStr(ReDimUI.App("navhome").State("homeShown") = 2)
    ReDimUI.AutoPump True
    SmokeNavigator = transcript
End Function
"""

SMOKE_GALLERY = """
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function SmokeGallery() As String
    Dim app As ReDimUI
    Dim transcript As String

    ReDimUI.AutoPump False
    BuildWidgetGallery
    Set app = GalleryApp()
    transcript = "components=" & app.ComponentCount

    ReDimUI.DispatchShape "rdm_gallery_primary"
    transcript = transcript & "|lastAction=" & app.State("lastAction")
    ReDimUI.DispatchShape "rdm_gallery_notify"
    transcript = transcript & "|toggleState=" & CStr(app.State("notifications"))
    transcript = transcript & "|inspectorLive=" & _
        CStr(InStr(CStr(app.State("inspector")), "notifications = True") > 0)

    ReDimUI.DispatchShape "rdm_gallery_region"
    Sleep 200
    ReDimUI.DispatchShape "rdm_gallery_region__opt3"
    transcript = transcript & "|regionPicked=" & app.State("region")

    Sleep 200
    ReDimUI.DispatchShape "rdm_gallery_busywork"
    transcript = transcript & "|asyncBusy=" & _
        CStr(app.Button("busywork").IsBusy)
    ReDimUI.PumpOnce
    transcript = transcript & "|asyncDone=" & _
        CStr(Not app.Button("busywork").IsBusy)
    ReDimUI.AutoPump True
    SmokeGallery = transcript
End Function
"""

SMOKE_SNAKE = """
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function SmokeSnake() As String
    Dim app As ReDimUI
    Dim headBefore As String
    Dim ticks As Long
    Dim transcript As String

    ReDimUI.AutoPump False
    BuildSnake
    Set app = SnakeApp()
    transcript = "components=" & app.ComponentCount

    NewGame
    transcript = transcript & "|active=" & CStr(IsGameActive())
    headBefore = HeadPosition()
    ' Paced at 150ms per move, so ticks only move the snake after the pace
    ' window elapses.
    For ticks = 1 To 3
        Sleep 170
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|headMoved=" & _
        CStr(HeadPosition() <> headBefore)

    KeyDown
    For ticks = 1 To 3
        Sleep 170
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|steered=" & CStr(HeadPosition() <> headBefore)
    transcript = transcript & "|stillActive=" & CStr(IsGameActive())
    transcript = transcript & "|scoreState=" & CStr(app.State("score") >= 0)
    UnhookKeys
    ReDimUI.AutoPump True
    SmokeSnake = transcript
End Function
"""


def open_demo(excel, demo_paths, name):
    excel.open_workbook(str(demo_paths[name]))


def test_mission_control_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Mission_Control.xlsm")
        result = excel.run_vba(SMOKE_MISSION, proc="SmokeMission", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 20
        assert facts["feed1Running"] == "True"
        assert facts["progressMoved"] == "True"
        assert facts["rowsMoved"] == "True"
        assert facts["launchDisabled"] == "True", (
            "launch must disable while any feed runs"
        )
        assert facts["canceledStatus"] == "Canceled"
        assert facts["startReEnabled"] == "True"
        assert facts["launchReEnabled"] == "True"
        assert facts["toastOnRail"] == "True", (
            "toast must sit on the rail beside the content"
        )
        assert facts["toastInkOnSurface"] == "True"
        assert facts["surfaceProtected"] == "True", (
            "the demo dashboard must ship with its surface protected"
        )


def test_widget_gallery_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Widget_Gallery.xlsm")
        result = excel.run_vba(SMOKE_GALLERY, proc="SmokeGallery", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 18
        assert facts["lastAction"] == "primary clicked"
        assert facts["toggleState"] == "True"
        assert facts["inspectorLive"] == "True"
        assert facts["regionPicked"] == "East"
        assert facts["asyncBusy"] == "True"
        assert facts["asyncDone"] == "True"


def test_navigator_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Navigator.xlsm")
        result = excel.run_vba(SMOKE_NAVIGATOR, proc="SmokeNavigator", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert facts["active"] == "navhome"
        assert facts["homeVisible"] == "True"
        assert facts["settingsHidden"] == "True"
        assert facts["homeShownOnce"] == "True"
        assert facts["homeProtected"] == "True"
        assert facts["homeTabActive"] == "True", (
            "the active window's tab must carry the primary fill"
        )
        assert facts["tabTitle"] == "Settings"
        assert facts["activeSettings"] == "True"
        assert facts["homeNowHidden"] == "True"
        assert facts["settingsTabActive"] == "True"
        assert facts["persistedDefault"] == "True"
        assert facts["backHome"] == "True"
        assert facts["homeShownTwice"] == "True"


def test_snake_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Snake.xlsm")
        result = excel.run_vba(SMOKE_SNAKE, proc="SmokeSnake", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 6
        assert facts["active"] == "True"
        assert facts["headMoved"] == "True"
        assert facts["steered"] == "True"
        assert facts["stillActive"] == "True"
        assert facts["scoreState"] == "True"
