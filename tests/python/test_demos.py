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

    ReDimUI.DispatchShape "rdm_mission_start1"
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

    Sleep 250
    ReDimUI.DispatchShape "rdm_mission_cancel1"
    ReDimUI.PumpOnce
    transcript = transcript & "|canceledStatus=" & app.State("feed1Status")
    transcript = transcript & "|startReEnabled=" & CStr(app.State("feed1Idle"))
    transcript = transcript & "|launchReEnabled=" & _
        CStr(app.Button("launch").IsEnabled)

    ' Toasts sit on a rail just outside the content's right edge, or
    ' inside the visible window when it is too narrow to hold one there.
    Dim toastShape As Shape
    Dim cardShape As Shape
    Dim railLeft As Double
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Set cardShape = app.Sheet.Shapes("rdm_mission_card1")
    Set toastShape = app.Sheet.Shapes("rdm_mission_toast_1")
    railLeft = cardShape.Left + cardShape.Width + 12
    app.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    If railLeft + toastShape.Width + 12 <= viewLeft + viewWidth Then
        transcript = transcript & "|toastOnRail=" & _
            CStr(Abs(toastShape.Left - railLeft) < 1)
    Else
        transcript = transcript & "|toastOnRail=" & _
            CStr(toastShape.Left >= viewLeft And _
                toastShape.Left + toastShape.Width <= viewLeft + viewWidth)
    End If
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
    transcript = transcript & "|defaultSeeded=" & _
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


SMOKE_POKEDEX = """
Public Function SmokePokeDex() As String
    Dim transcript As String

    ReDimUI.AutoPump False
    BuildPokeDex
    transcript = "active=" & ReDimUI.ActiveWindowId
    transcript = transcript & "|browseVisible=" & _
        CStr(ThisWorkbook.Worksheets("DexBrowse").Visible = xlSheetVisible)
    transcript = transcript & "|teamHidden=" & _
        CStr(ThisWorkbook.Worksheets("DexTeam").Visible = xlSheetVeryHidden)
    transcript = transcript & "|protected=" & _
        CStr(ThisWorkbook.Worksheets("DexBrowse").ProtectContents)
    transcript = transcript & "|keyShapes=" & _
        CStr(ShapeThere("DexBrowse", "rdm_dexbrowse_species") And _
            ShapeThere("DexBrowse", "rdm_dexbrowse_sprite") And _
            ShapeThere("DexBrowse", "rdm_dexbrowse_statb6") And _
            ShapeThere("DexBrowse", "rdm_dexbrowse_nvb_dexteam"))
    transcript = transcript & "|fetchKicked=" & _
        CStr(ReDimUI.App("dexbrowse").Async("names").IsRunning)

    ReDimUI.DispatchShape "rdm_dexbrowse_nvb_dexteam"
    transcript = transcript & "|teamActive=" & _
        CStr(ReDimUI.ActiveWindowId = "dexteam")
    transcript = transcript & "|teamShapes=" & _
        CStr(ShapeThere("DexTeam", "rdm_dexteam_team__mvr") And _
            ShapeThere("DexTeam", "rdm_dexteam_prefs__mb") And _
            ShapeThere("DexTeam", "rdm_dexteam_strategy"))
    ReDimUI.DispatchShape "rdm_dexteam_nvb_dextrainer"
    transcript = transcript & "|trainerActive=" & _
        CStr(ReDimUI.ActiveWindowId = "dextrainer")
    transcript = transcript & "|trainerShapes=" & _
        CStr(ShapeThere("DexTrainer", "rdm_dextrainer_darkmode") And _
            ShapeThere("DexTrainer", "rdm_dextrainer_starter") And _
            ShapeThere("DexTrainer", "rdm_dextrainer_reset"))
    transcript = transcript & "|themedPrimary=" & _
        CStr(ReDimUI.App("dextrainer").Theme.PrimaryColor = RGB(214, 55, 46))

    ReDimUI.App("dextrainer").SetState "darkMode", True
    ApplyThemeChoice
    transcript = transcript & "|nightCanvas=" & _
        CStr(ThisWorkbook.Worksheets("DexBrowse").Cells(1, 1).Interior.Color _
            = ReDimUI.App("dexbrowse").Theme.CanvasColor)
    transcript = transcript & "|nightTrack=" & _
        CStr(ThisWorkbook.Worksheets("DexBrowse").Shapes( _
            "rdm_dexbrowse_statb1").Fill.ForeColor.RGB _
            = ReDimUI.App("dexbrowse").Theme.MutedColor)
    transcript = transcript & "|nightIsDark=" & _
        CStr(ReDimUI.App("dexbrowse").Theme.PrimaryColor <> RGB(214, 55, 46))

    Dim probe As ROneCOne
    Dim spritesNode As ROneCOne
    Dim raw As String

    ' Quotes are written as apostrophes and swapped in, so the fixture
    ' reads cleanly and carries no doubled-quote runs.
    raw = Replace("{'id':25,'moves':[{'m':1}],'name':'pikachu'," & _
        "'sprites':{'front_default':'u.png','versions':{'z':1}}}", _
        "'", Chr$(34))
    Set probe = ROneCOne.Json.DeserializeOnly(raw, _
        Array("$.id", "$.name", "$.sprites.front_default"))
    Set spritesNode = probe.Item("sprites")
    transcript = transcript & "|partialName=" & _
        CStr(CStr(probe.Item("name")) = "pikachu")
    transcript = transcript & "|partialSprite=" & _
        CStr(CStr(spritesNode.Item("front_default")) = "u.png")
    transcript = transcript & "|partialSkipped=" & _
        CStr(Not probe.ContainsKey("moves"))
    transcript = transcript & "|partialPruned=" & _
        CStr(Not spritesNode.ContainsKey("versions"))

    ' The library surfaces the dex leans on in place of hand-rolled
    ' equivalents: a typed cache, temp paths, and file probing.
    Dim cache As ROneCOne
    Dim tempFile As String
    Set cache = ROneCOne.DictionaryOf(vbString, vbObject)
    cache.Add "bulbasaur", probe
    transcript = transcript & "|cacheHit=" & _
        CStr(cache.ContainsKey("bulbasaur") And Not cache.ContainsKey("mew"))
    tempFile = ROneCOne.Path.Combine( _
        ROneCOne.Path.GetTempPath(), "redex_probe_absent.png")
    transcript = transcript & "|tempRooted=" & _
        CStr(InStr(tempFile, ROneCOne.Path.GetTempPath()) = 1)
    transcript = transcript & "|absentFile=" & _
        CStr(Not ROneCOne.File.Exists(tempFile))
    transcript = transcript & "|spriteIdle=" & _
        CStr(Not ReDimUI.App("dexbrowse").Async("sprite").IsRunning)
    RdxReleaseKeys
    RdxStopPump
    ReDimUI.AutoPump True
    SmokePokeDex = transcript
End Function

Private Function ShapeThere( _
    ByVal sheetName As String, ByVal shapeName As String) As Boolean
    Dim target As Shape
    On Error Resume Next
    Set target = ThisWorkbook.Worksheets(sheetName).Shapes(shapeName)
    On Error GoTo 0
    ShapeThere = Not target Is Nothing
End Function
"""


SMOKE_EXPENSES = """
Public Function SmokeExpenses() As String
    Dim app As ReDimUI
    Dim transcript As String
    Dim rowsBefore As Long

    ReDimUI.AutoPump False
    BuildExpenseTracker
    Set app = ExpenseApp()
    transcript = "components=" & app.ComponentCount
    rowsBefore = app.Table("list").RowCount
    transcript = transcript & "|seeded=" & CStr(rowsBefore >= 30)
    transcript = transcript & "|trend=" & _
        Left$(app.Sheet.Shapes("rdm_expenses_trend").AlternativeText, 15)
    transcript = transcript & "|dataHidden=" & _
        CStr(ThisWorkbook.Worksheets("ExpenseData").Visible = xlSheetHidden)

    HandleAdd
    transcript = transcript & "|blocked=" & _
        CStr(app.Table("list").RowCount = rowsBefore) & "/" & _
        app.TextInput("amount").ValidationError
    app.TextInput("amount").InputValue = "42.5"
    app.TextInput("note").InputValue = "Smoke lunch"
    app.SelectBox("category").Value 2
    HandleAdd
    transcript = transcript & "|added=" & _
        CStr(app.Table("list").RowCount = rowsBefore + 1) & "/" & _
        app.TextInput("amount").InputValue & "/" & _
        CStr(ReDimUI.IsComponentFocused("expenses", "amount"))

    app.Table("list").FilterRows "smoke lunch"
    transcript = transcript & "|found=" & app.Table("list").ShownRowCount
    HandleExport
    With ThisWorkbook.Worksheets("ExpenseExport")
        transcript = transcript & "|exported=" & .Range("A1").Value & "," & _
            .Range("B2").Value & "," & .Range("D2").Value & "," & _
            CStr(IsEmpty(.Range("A3").Value))
    End With
    transcript = transcript & "|stayed=" & CStr(ActiveSheet.Name = "Expenses")
    app.Table("list").FilterRows ""

    RequestDelete rowsBefore + 1
    ReDimUI.DispatchShape "rdm_expenses_mdl_ok"
    transcript = transcript & "|deleted=" & CStr(app.Table("list").RowCount = rowsBefore)
    UndoDelete
    transcript = transcript & "|restored=" & _
        CStr(app.Table("list").RowCount = rowsBefore + 1)

    FlipDarkMode
    transcript = transcript & "|dark=" & CStr(app.Theme.SurfaceColor = _
        ReDimUI.ThemeDark.SurfaceColor And app.Toggle("dark").IsChecked)
    transcript = transcript & "|protected=" & CStr(app.Sheet.ProtectContents)
    RdxReleaseKeys
    RdxStopPump
    ReDimUI.AutoPump True
    SmokeExpenses = transcript
End Function
"""


SMOKE_WHATSNEW = """
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function SmokeWhatsNew() As String
    Dim app As ReDimUI
    Dim transcript As String
    Dim ticks As Long

    ReDimUI.AutoPump False
    BuildWhatsNew
    Set app = WhatsNewApp()
    transcript = "components=" & app.ComponentCount

    ' Code sets the key two controls write, and both follow.
    ReDimUI.DispatchShape "rdm_whatsnew_setLarge"
    transcript = transcript & "|sizeFollows=" & app.SelectBox("sizeSelect").SelectedText & _
        "/" & app.RadioGroup("sizeRadio").SelectedText

    ' One handler serves the buttons by their Tag, and the stepper follows.
    Sleep 200
    ReDimUI.DispatchShape "rdm_whatsnew_add5"
    Sleep 200
    ReDimUI.DispatchShape "rdm_whatsnew_add10"
    transcript = transcript & "|count=" & app.State("count") & "/" & _
        app.Stepper("count").CurrentValue

    HandleFind
    transcript = transcript & "|found=" & app.SelectBox("fruit").SelectedText

    app.SetState "themeName", "Ocean"
    transcript = transcript & "|ocean=" & CStr(app.Theme.PrimaryColor = RGB(0, 99, 140))

    app.SetState "section", "Async and errors"
    transcript = transcript & "|section=" & app.Tabs("sections").CurrentValue

    ' The failing handler is left out: an error raised inside a handler
    ' escapes a harness call whatever traps it, as TestReDimAsync notes.
    Sleep 200
    ReDimUI.DispatchShape "rdm_whatsnew_roll"
    For ticks = 1 To 50
        If Not app.Async("dieRoll").IsRunning Then Exit For
        Sleep 60
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|rolled=" & CStr(CStr(app.State("roll")) Like "Rolled a [1-6] *")

    Sleep 200
    ReDimUI.DispatchShape "rdm_whatsnew_fillA"
    Sleep 200
    ReDimUI.DispatchShape "rdm_whatsnew_fillB"
    For ticks = 1 To 120
        If Not app.Job("jobA").IsRunning And Not app.Job("jobB").IsRunning Then Exit For
        Sleep 60
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|filled=" & app.State("jobA") & "/" & app.State("jobB")
    transcript = transcript & "|protected=" & CStr(app.Sheet.ProtectContents)
    RdxReleaseKeys
    RdxStopPump
    ReDimUI.AutoPump True
    SmokeWhatsNew = transcript
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
            "toast must sit on the rail beside the content, or inside a "
            "window too narrow for the rail"
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
        assert facts["defaultSeeded"] == "True"
        assert facts["backHome"] == "True"
        assert facts["homeShownTwice"] == "True"


def test_pokedex_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_ReDex.xlsm")
        result = excel.run_vba(SMOKE_POKEDEX, proc="SmokePokeDex", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert facts["active"] == "dexbrowse"
        assert facts["browseVisible"] == "True"
        assert facts["teamHidden"] == "True"
        assert facts["protected"] == "True"
        assert facts["keyShapes"] == "True", (
            "the browse window must carry the combo, sprite image, stat"
            " bars, and nav tabs"
        )
        assert facts["fetchKicked"] == "True", (
            "showing the browse window must start the species fetch op"
        )
        assert facts["teamActive"] == "True"
        assert facts["teamShapes"] == "True", (
            "the team window must carry the transfer list, checklist"
            " header, and notes field"
        )
        assert facts["trainerActive"] == "True"
        assert facts["trainerShapes"] == "True"
        assert facts["themedPrimary"] == "True", (
            "the custom Pokedex theme must drive the whole app"
        )
        assert facts["nightCanvas"] == "True", (
            "night mode must repaint the browse canvas background"
        )
        assert facts["nightTrack"] == "True", (
            "night mode must retint the stat bar tracks"
        )
        assert facts["nightIsDark"] == "True"
        assert facts["partialName"] == "True"
        assert facts["partialSprite"] == "True", (
            "a nested allowlist path must stay navigable as a subtree"
        )
        assert facts["partialSkipped"] == "True", (
            "an unrequested top-level member must never be materialized"
        )
        assert facts["partialPruned"] == "True", (
            "a kept parent must carry only the requested child"
        )
        assert facts["cacheHit"] == "True", (
            "the typed dictionary must answer membership directly"
        )
        assert facts["tempRooted"] == "True"
        assert facts["absentFile"] == "True"
        assert facts["spriteIdle"] == "True", (
            "no sprite download may be in flight before any species loads"
        )


def test_expense_tracker_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Expense_Tracker.xlsm")
        result = excel.run_vba(SMOKE_EXPENSES, proc="SmokeExpenses", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 20
        assert facts["seeded"] == "True", "a first build seeds six months of samples"
        assert facts["trend"] == "Trend, 6 values"
        assert facts["dataHidden"] == "True"
        assert facts["blocked"] == "True/Required", (
            "an empty amount holds the form back with its message"
        )
        assert facts["added"] == "True//True", (
            "an add clears the amount and puts focus back on it"
        )
        assert facts["found"] == "1"
        assert facts["exported"] == "Date,Dining,Smoke lunch,True", (
            "the export writes the filtered view under its header"
        )
        assert facts["stayed"] == "True", "the export leaves the tracker in front"
        assert facts["deleted"] == "True"
        assert facts["restored"] == "True"
        assert facts["dark"] == "True", "the palette's command flips the theme and the switch"
        assert facts["protected"] == "True"


def test_whats_new_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Whats_New.xlsm")
        result = excel.run_vba(SMOKE_WHATSNEW, proc="SmokeWhatsNew", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 60
        assert facts["sizeFollows"] == "Large/Large", "both controls follow the key code sets"
        assert facts["count"] == "15/15", "one handler adds each button's Tag"
        assert facts["found"] == "Mangosteen", "ItemPosition finds an item in any case"
        assert facts["ocean"] == "True", "a theme built with the With builders"
        assert facts["section"] == "5", "the tab strip follows its key"
        assert facts["rolled"] == "True", "the done handler reads the task's result"
        assert facts["filled"] == "100/100", "two jobs share one step by their Tag"
        assert facts["protected"] == "True"


def test_snake_smoke(demo_paths):
    with ExcelSession() as excel:
        open_demo(excel, demo_paths, "ReDim_Snake.xlsm")
        result = excel.run_vba(SMOKE_SNAKE, proc="SmokeSnake", timeout=120)
        assert result.outcome == "passed", result.error
        facts = dict(t.split("=", 1) for t in result.value.split("|"))
        assert int(facts["components"]) >= 5
        assert facts["active"] == "True"
        assert facts["headMoved"] == "True"
        assert facts["steered"] == "True"
        assert facts["stillActive"] == "True"
        assert facts["scoreState"] == "True"
