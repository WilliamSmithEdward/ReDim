Attribute VB_Name = "MissionControl"
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
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim feed As Long

    Set host = ThisWorkbook.Worksheets(1)
    Set app = ReDimUI.Mount(host, APP_ID)
    app.PrepareCanvas

    app.Label("title").AtRect(24, 16, 360, 30).Text("Mission Control") _
        .FontSize(20).Bold
    app.Label("subtitle").AtRect(24, 48, 420, 18) _
        .Text("Three async feeds, one responsive workbook. Built with ReDim.")
    app.Spinner("busy").AtRect 400, 18, 26, 26
    app.Toggle("dark").AtRect(452, 22, 44, 22).WritesTo("darkMode") _
        .OnChange "MissionControl.HandleThemeToggle"
    app.Label("darklbl").AtRect(502, 24, 80, 18).Text("Dark mode")

    app.Button("launch").AtRect(24, 78, 120, 32).Text("Launch all feeds") _
        .Primary.BindEnabled("anyRunning", True) _
        .OnClick "MissionControl.HandleLaunchAll"
    app.Button("reset").AtRect(152, 78, 90, 32).Text("Reset").Danger _
        .OnClick "MissionControl.HandleResetRequest"

    For feed = 1 To 3
        BuildFeedPanel app, feed
    Next feed

    app.Card("kpiRows").AtRect(24, 320, 160, 84).Text("Rows loaded")
    app.Label("kpiRowsVal").AtRect(36, 352, 130, 40).BindText("rowsLoaded") _
        .FontSize(24).Bold
    app.Card("kpiDone").AtRect(196, 320, 160, 84).Text("Feeds complete")
    app.Label("kpiDoneVal").AtRect(208, 352, 130, 40) _
        .BindText("feedsDone", "{0} of 3").FontSize(24).Bold

    app.SetState "darkMode", False
    app.SetState "rowsLoaded", 0
    app.SetState "feedsDone", "0"
    app.SetState "anyRunning", False
    ResetFeedState app, 1
    ResetFeedState app, 2
    ResetFeedState app, 3
    app.Spinner("busy").BindVisible "anyRunning"
    app.Render
End Sub

Private Sub BuildFeedPanel(ByVal app As ReDimUI, ByVal feed As Long)
    Dim panelTop As Double
    Dim key As String

    panelTop = 126 + (feed - 1) * 62
    key = FeedKey(feed)
    app.Card("card" & feed).AtRect(24, panelTop, 560, 54).Text(vbNullString)
    app.Label("name" & feed).AtRect(36, panelTop + 8, 110, 18) _
        .Text(FeedName(feed)).Bold
    app.Label("stat" & feed).AtRect(36, panelTop + 28, 130, 16) _
        .BindText key & "Status"
    app.ProgressBar("prg" & feed).AtRect(180, panelTop + 20, 240, 12) _
        .BindValue key & "Pct"
    app.Button("start" & feed).AtRect(436, panelTop + 12, 64, 28) _
        .Text("Start").Primary.BindEnabled(key & "Idle") _
        .OnClick "MissionControl.HandleStartFeed" & feed
    app.Button("cancel" & feed).AtRect(508, panelTop + 12, 64, 28) _
        .Text("Cancel").Secondary.BindEnabled(key & "Running") _
        .OnClick "MissionControl.HandleCancelFeed" & feed
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

Private Sub ResetFeedState(ByVal app As ReDimUI, ByVal feed As Long)
    gFeedPct(feed) = 0
    app.SetState FeedKey(feed) & "Pct", 0
    app.SetState FeedKey(feed) & "Status", "Idle"
    app.SetState FeedKey(feed) & "Idle", True
    app.SetState FeedKey(feed) & "Running", False
End Sub

' ---------------------------------------------------------------
' Handlers
' ---------------------------------------------------------------

Public Sub HandleThemeToggle()
    Dim app As ReDimUI

    Set app = MissionApp()
    If CBool(app.State("darkMode")) Then
        app.SetTheme ReDimUI.ThemeDark
    Else
        app.SetTheme ReDimUI.ThemeLight
    End If
    app.PrepareCanvas
End Sub

Public Sub HandleLaunchAll()
    StartFeed 1
    StartFeed 2
    StartFeed 3
    MissionApp().Toast "All feeds launched.", 2500
End Sub

Public Sub HandleStartFeed1()
    StartFeed 1
End Sub

Public Sub HandleStartFeed2()
    StartFeed 2
End Sub

Public Sub HandleStartFeed3()
    StartFeed 3
End Sub

Public Sub HandleCancelFeed1()
    MissionApp().CancelJob "job1"
End Sub

Public Sub HandleCancelFeed2()
    MissionApp().CancelJob "job2"
End Sub

Public Sub HandleCancelFeed3()
    MissionApp().CancelJob "job3"
End Sub

Public Sub HandleResetRequest()
    MissionApp().Confirm "Reset dashboard", _
        "Stop every feed and clear all progress?", _
        "MissionControl.HandleResetConfirmed"
End Sub

Public Sub HandleResetConfirmed()
    Dim app As ReDimUI
    Dim feed As Long

    Set app = MissionApp()
    For feed = 1 To 3
        On Error Resume Next
        app.CancelJob "job" & feed
        On Error GoTo 0
        ResetFeedState app, feed
    Next feed
    gRowsLoaded = 0
    gFeedsDone = 0
    app.SetState "rowsLoaded", 0
    app.SetState "feedsDone", "0"
    app.SetState "anyRunning", False
    app.Toast "Dashboard reset.", 2000
End Sub

Private Sub StartFeed(ByVal feed As Long)
    Dim app As ReDimUI
    Dim key As String

    Set app = MissionApp()
    key = FeedKey(feed)
    If app.Job("job" & feed).JobIsRunning Then Exit Sub
    gFeedPct(feed) = 0
    app.SetState key & "Pct", 0
    app.SetState key & "Status", "Loading"
    app.SetState key & "Idle", False
    app.SetState key & "Running", True
    app.SetState "anyRunning", True
    app.Job("job" & feed).Steps("MissionControl.FeedStep" & feed) _
        .PacedMs(50 + feed * 30) _
        .JobOnDone("MissionControl.FeedDone" & feed) _
        .JobOnCancel "MissionControl.FeedCanceled" & feed
    app.Job("job" & feed).StartJob
End Sub

' One paced step is one arriving chunk: an instant progress increment with
' no blocking work, so the pump's duty cycle stays negligible.
Private Function FeedStep(ByVal feed As Long) As Boolean
    Dim app As ReDimUI

    Set app = MissionApp()
    gFeedPct(feed) = gFeedPct(feed) + 1 + (feed Mod 3)
    gRowsLoaded = gRowsLoaded + 25 + feed * 5
    app.SetState FeedKey(feed) & "Pct", gFeedPct(feed)
    app.SetState "rowsLoaded", gRowsLoaded
    FeedStep = (gFeedPct(feed) >= 100)
End Function

Public Function FeedStep1() As Boolean
    FeedStep1 = FeedStep(1)
End Function

Public Function FeedStep2() As Boolean
    FeedStep2 = FeedStep(2)
End Function

Public Function FeedStep3() As Boolean
    FeedStep3 = FeedStep(3)
End Function

Private Sub FeedFinished(ByVal feed As Long, ByVal finalStatus As String)
    Dim app As ReDimUI
    Dim anyRunning As Boolean
    Dim other As Long

    Set app = MissionApp()
    app.SetState FeedKey(feed) & "Status", finalStatus
    app.SetState FeedKey(feed) & "Idle", True
    app.SetState FeedKey(feed) & "Running", False
    For other = 1 To 3
        If app.Job("job" & other).JobIsRunning Then anyRunning = True
    Next other
    app.SetState "anyRunning", anyRunning
    If finalStatus = "Complete" Then
        gFeedsDone = gFeedsDone + 1
        app.SetState "feedsDone", CStr(gFeedsDone)
        app.SetState FeedKey(feed) & "Pct", 100
        app.Toast FeedName(feed) & " finished.", 2500
    Else
        app.Toast FeedName(feed) & " canceled.", 2500
    End If
End Sub

Public Sub FeedDone1()
    FeedFinished 1, "Complete"
End Sub

Public Sub FeedDone2()
    FeedFinished 2, "Complete"
End Sub

Public Sub FeedDone3()
    FeedFinished 3, "Complete"
End Sub

Public Sub FeedCanceled1()
    FeedFinished 1, "Canceled"
End Sub

Public Sub FeedCanceled2()
    FeedFinished 2, "Canceled"
End Sub

Public Sub FeedCanceled3()
    FeedFinished 3, "Canceled"
End Sub
