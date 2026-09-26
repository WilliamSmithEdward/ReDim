Attribute VB_Name = "MissionControl"
' ReDim 1.0.2 (2026-09-23)
' https://github.com/WilliamSmithEdward/ReDim
'
' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

' Mission Control: the ReDim flagship demo. Three simulated data feeds run as
' paced jobs with live progress bars, individual cancel buttons, completion
' toasts, a global activity spinner, KPI cards, a dark mode toggle, and a
' shapes-based confirm modal. Excel stays fully interactive while feeds run.
' Steps never block: each paced step is one instant progress increment, so
' the pump's duty cycle stays tiny and the cursor stays calm.

Private Const APP_ID As String = "mission"

Private gFeedPct(1 To 3) As Double
Private gRowsLoaded As Long
Private gFeedsDone As Long

Public Sub Auto_Open()
    BuildMissionControl
End Sub

Public Function MissionApp() As ReDimUI
    Set MissionApp = ReDimUI.App(APP_ID)
End Function

Public Sub BuildMissionControl()
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim feed As Long

    Set host = ThisWorkbook.Worksheets(1)
    Set ui = ReDimUI.Mount(host, APP_ID)
    ' UserInterfaceOnly protection does not survive reopen, so builds
    ' unprotect first and re-protect at the end.
    ui.ProtectSurface False
    ui.PrepareCanvas

    ui.Label("title").AtRect(24, 16, 360, 30).Text("Mission Control") _
        .FontSize(20).Bold
    ui.Label("subtitle").AtRect(24, 48, 420, 18) _
        .Text("Three async feeds, one responsive workbook. Built with ReDim.")
    ui.Spinner("busy").AtRect 400, 18, 26, 26
    ui.Toggle("dark").AtRect(452, 22, 44, 22).Text("Dark mode").WritesTo("darkMode") _
        .OnChange "MissionControl.HandleThemeToggle"

    ui.Button("launch").AtRect(24, 78, 120, 32).Text("Launch all feeds") _
        .Primary.BindEnabled("anyRunning", True) _
        .OnClick "MissionControl.HandleLaunchAll"
    ui.Button("reset").AtRect(152, 78, 90, 32).Text("Reset").Danger _
        .OnClick "MissionControl.HandleResetRequest"

    For feed = 1 To 3
        BuildFeedPanel ui, feed
    Next feed

    ui.Card("kpiRows").AtRect(24, 320, 160, 84).Text("Rows loaded")
    ui.Label("kpiRowsVal").AtRect(36, 352, 130, 40).BindText("rowsLoaded") _
        .FontSize(24).Bold
    ui.Card("kpiDone").AtRect(196, 320, 160, 84).Text("Feeds complete")
    ui.Label("kpiDoneVal").AtRect(208, 352, 130, 40) _
        .BindText("feedsDone", "{0} of 3").FontSize(24).Bold

    ui.SetState "darkMode", False
    ui.SetState "rowsLoaded", 0
    ui.SetState "feedsDone", "0"
    ui.SetState "anyRunning", False
    ResetFeedState ui, 1
    ResetFeedState ui, 2
    ResetFeedState ui, 3
    ui.Spinner("busy").BindVisible "anyRunning"
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub BuildFeedPanel(ByVal ui As ReDimUI, ByVal feed As Long)
    Dim panelTop As Double
    Dim keyPrefix As String

    panelTop = 126 + (feed - 1) * 62
    keyPrefix = FeedKey(feed)
    ui.Card("card" & feed).AtRect(24, panelTop, 560, 54).Text(vbNullString)
    ui.Label("name" & feed).AtRect(36, panelTop + 8, 110, 18) _
        .Text(FeedName(feed)).Bold
    ui.Label("stat" & feed).AtRect(36, panelTop + 28, 130, 16) _
        .BindText keyPrefix & "Status"
    ui.ProgressBar("prg" & feed).AtRect(180, panelTop + 20, 240, 12) _
        .BindValue keyPrefix & "Pct"
    ' Each feed's buttons carry its number, so one handler serves all three.
    ui.Button("start" & feed).AtRect(436, panelTop + 12, 64, 28) _
        .Text("Start").Primary.BindEnabled(keyPrefix & "Idle") _
        .OnClick("MissionControl.HandleStartFeed").Tag feed
    ui.Button("cancel" & feed).AtRect(508, panelTop + 12, 64, 28) _
        .Text("Cancel").Secondary.BindEnabled(keyPrefix & "Running") _
        .OnClick("MissionControl.HandleCancelFeed").Tag feed
End Sub

Private Function FeedKey(ByVal feed As Long) As String
    FeedKey = "feed" & feed
End Function

Private Function FeedName(ByVal feed As Long) As String
    Select Case feed
        Case 1
            FeedName = "Alpha telemetry"
        Case 2
            FeedName = "Beta ledger"
        Case Else
            FeedName = "Gamma sensors"
    End Select
End Function

Private Sub ResetFeedState(ByVal ui As ReDimUI, ByVal feed As Long)
    gFeedPct(feed) = 0
    ui.SetState FeedKey(feed) & "Pct", 0
    ui.SetState FeedKey(feed) & "Status", "Idle"
    ui.SetState FeedKey(feed) & "Idle", True
    ui.SetState FeedKey(feed) & "Running", False
End Sub

' ---------------------------------------------------------------
' Handlers
' ---------------------------------------------------------------

Public Sub HandleThemeToggle()
    Dim ui As ReDimUI

    Set ui = MissionApp()
    If CBool(ui.State("darkMode")) Then
        ui.SetTheme ReDimUI.ThemeDark
    Else
        ui.SetTheme ReDimUI.ThemeLight
    End If
End Sub

Public Sub HandleLaunchAll()
    StartFeed 1
    StartFeed 2
    StartFeed 3
    MissionApp().Toast "All feeds launched.", 2500
End Sub

Public Sub HandleStartFeed()
    StartFeed CLng(ReDimUI.Sender.TagValue)
End Sub

Public Sub HandleCancelFeed()
    MissionApp().CancelJob "job" & ReDimUI.Sender.TagValue
End Sub

Public Sub HandleResetRequest()
    MissionApp().Confirm "Reset dashboard", _
        "Stop every feed and clear all progress?", _
        "MissionControl.HandleResetConfirmed"
End Sub

Public Sub HandleResetConfirmed()
    Dim ui As ReDimUI
    Dim feed As Long

    Set ui = MissionApp()
    For feed = 1 To 3
        ui.CancelJob "job" & feed
        ResetFeedState ui, feed
    Next feed
    gRowsLoaded = 0
    gFeedsDone = 0
    ui.SetState "rowsLoaded", 0
    ui.SetState "feedsDone", "0"
    ui.SetState "anyRunning", False
    ui.Toast "Dashboard reset.", 2000
End Sub

Private Sub StartFeed(ByVal feed As Long)
    Dim ui As ReDimUI
    Dim keyPrefix As String

    Set ui = MissionApp()
    keyPrefix = FeedKey(feed)
    If ui.Job("job" & feed).IsRunning Then Exit Sub
    gFeedPct(feed) = 0
    ui.SetState keyPrefix & "Pct", 0
    ui.SetState keyPrefix & "Status", "Loading"
    ui.SetState keyPrefix & "Idle", False
    ui.SetState keyPrefix & "Running", True
    ui.SetState "anyRunning", True
    ' The job carries its feed's number too: its step and outcome handlers
    ' read it from ReDimUI.Sender.
    ui.Job("job" & feed).Tag(feed).Steps("MissionControl.FeedStep") _
        .PacedMs(50 + feed * 30) _
        .JobOnDone("MissionControl.FeedDone") _
        .JobOnCancel "MissionControl.FeedCanceled"
    ui.Job("job" & feed).StartJob
End Sub

' One paced step is one arriving chunk: an instant progress increment with
' no blocking work, so the pump's duty cycle stays negligible.
Public Function FeedStep() As Boolean
    Dim ui As ReDimUI
    Dim feed As Long

    Set ui = MissionApp()
    feed = CLng(ReDimUI.Sender.TagValue)
    gFeedPct(feed) = gFeedPct(feed) + 1 + (feed Mod 3)
    gRowsLoaded = gRowsLoaded + 25 + feed * 5
    ui.SetState FeedKey(feed) & "Pct", gFeedPct(feed)
    ui.SetState "rowsLoaded", gRowsLoaded
    FeedStep = (gFeedPct(feed) >= 100)
End Function

Private Sub FeedFinished(ByVal feed As Long, ByVal finalStatus As String)
    Dim ui As ReDimUI
    Dim anyRunning As Boolean
    Dim otherFeed As Long

    Set ui = MissionApp()
    ui.SetState FeedKey(feed) & "Status", finalStatus
    ui.SetState FeedKey(feed) & "Idle", True
    ui.SetState FeedKey(feed) & "Running", False
    For otherFeed = 1 To 3
        If ui.Job("job" & otherFeed).IsRunning Then anyRunning = True
    Next otherFeed
    ui.SetState "anyRunning", anyRunning
    If finalStatus = "Complete" Then
        gFeedsDone = gFeedsDone + 1
        ui.SetState "feedsDone", CStr(gFeedsDone)
        ui.SetState FeedKey(feed) & "Pct", 100
        ui.Toast FeedName(feed) & " finished.", 2500
    Else
        ui.Toast FeedName(feed) & " canceled.", 2500
    End If
End Sub

Public Sub FeedDone()
    FeedFinished CLng(ReDimUI.Sender.TagValue), "Complete"
End Sub

Public Sub FeedCanceled()
    FeedFinished CLng(ReDimUI.Sender.TagValue), "Canceled"
End Sub
