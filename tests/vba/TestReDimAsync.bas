Attribute VB_Name = "TestReDimAsync"
Option Explicit

' Live scenarios for the ReDim async engine: ops, the OnClickAsync path,
' transport tasks across ticks, cancellation, chunked jobs, and one real
' SetTimer end-to-end run. Deterministic scenarios turn AutoPump off and
' drive PumpOnce; the end-to-end scenario uses the real timer and proves it
' disarms itself before returning.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private gWorkRan As Long
Private gDoneRan As Long
Private gCancelRan As Long
Private gDuringWorkBusy As Boolean
Private gJobCounter As Long

Private Function NewCanvas() As Worksheet
    Set NewCanvas = ActiveWorkbook.Worksheets.Add
End Function

Public Sub RecordWork()
    gWorkRan = gWorkRan + 1
End Sub

Public Sub RecordWorkCapturingBusy()
    gWorkRan = gWorkRan + 1
    gDuringWorkBusy = ReDimUI.App("async1").Button("btn").IsBusy
End Sub

Public Sub RecordDone()
    gDoneRan = gDoneRan + 1
End Sub

Public Sub RecordCancel()
    gCancelRan = gCancelRan + 1
End Sub

Public Function JobStepSmall() As Boolean
    gJobCounter = gJobCounter + 1
    Sleep 3
    JobStepSmall = (gJobCounter >= 5)
End Function

Public Function JobStepEndless() As Boolean
    gJobCounter = gJobCounter + 1
    Sleep 3
    JobStepEndless = (gJobCounter >= 1000)
End Function

Public Function TestAsyncOpLifecycle() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gWorkRan = 0
    gDoneRan = 0
    gDuringWorkBusy = False
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async1")
    app.Button("btn").At("B2:C3").Text("Run").Primary
    app.Spinner("spn").At("E2")
    app.Label("status").At("B5:E5").BindText "opStatus", "Op: {0}"
    app.Async("op1").RunsProc("TestReDimAsync.RecordWorkCapturingBusy") _
        .Disables("btn").ShowsSpinner("spn") _
        .OnDone("TestReDimAsync.RecordDone").TracksState "opStatus"
    app.Render

    app.Async("op1").Start
    transcript = "statusAfterStart=" & app.State("opStatus")
    transcript = transcript & "|busyText=" & _
        host.Shapes("rdm_async1_btn").TextFrame2.TextRange.Text
    transcript = transcript & "|spinnerShown=" & _
        CStr(host.Shapes("rdm_async1_spn").Visible = msoTrue)
    transcript = transcript & "|workBeforeTick=" & gWorkRan
    ReDimUI.DispatchShape "rdm_async1_btn"
    transcript = transcript & "|clickWhileBusyIgnored=" & CStr(gWorkRan = 0)

    ReDimUI.PumpOnce
    transcript = transcript & "|workAfterTick=" & gWorkRan
    transcript = transcript & "|busyDuringWork=" & CStr(gDuringWorkBusy)
    transcript = transcript & "|doneRan=" & gDoneRan
    transcript = transcript & "|statusAfterTick=" & app.State("opStatus")
    transcript = transcript & "|restoredText=" & _
        host.Shapes("rdm_async1_btn").TextFrame2.TextRange.Text
    transcript = transcript & "|spinnerHidden=" & _
        CStr(host.Shapes("rdm_async1_spn").Visible = msoFalse)
    transcript = transcript & "|labelText=" & _
        host.Shapes("rdm_async1_status").TextFrame2.TextRange.Text
    ReDimUI.AutoPump True
    TestAsyncOpLifecycle = transcript
End Function

Public Function TestOnClickAsyncSugar() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gWorkRan = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async2")
    app.Button("btnrun").At("B2:C3").Text("Run").Primary _
        .OnClickAsync "TestReDimAsync.RecordWork"
    app.Render

    ReDimUI.DispatchShape "rdm_async2_btnrun"
    transcript = "busyAfterClick=" & CStr(app.Button("btnrun").IsBusy)
    transcript = transcript & "|textAfterClick=" & _
        host.Shapes("rdm_async2_btnrun").TextFrame2.TextRange.Text
    Sleep 200
    ReDimUI.DispatchShape "rdm_async2_btnrun"
    transcript = transcript & "|workBeforeTick=" & gWorkRan

    ReDimUI.PumpOnce
    transcript = transcript & "|workAfterTick=" & gWorkRan
    transcript = transcript & "|busyAfterTick=" & CStr(app.Button("btnrun").IsBusy)
    transcript = transcript & "|restoredText=" & _
        host.Shapes("rdm_async2_btnrun").TextFrame2.TextRange.Text
    Sleep 200
    ReDimUI.DispatchShape "rdm_async2_btnrun"
    ReDimUI.PumpOnce
    transcript = transcript & "|secondRunWorks=" & CStr(gWorkRan = 2)
    ReDimUI.AutoPump True
    TestOnClickAsyncSugar = transcript
End Function

Public Function TestTransportOpAcrossTicks() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async3")
    app.Async("wait").RunsTask(ROneCOne.Task.Delay(150)).TracksState "st"
    app.Render
    app.Async("wait").Start

    ReDimUI.PumpOnce
    transcript = "midStatus=" & app.State("st")
    Sleep 220
    ReDimUI.PumpOnce
    transcript = transcript & "|finalStatus=" & app.State("st")
    ReDimUI.AutoPump True
    TestTransportOpAcrossTicks = transcript
End Function

Public Function TestCancellation() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim opValue As ReDimUI
    Dim transcript As String

    Set host = NewCanvas()
    gWorkRan = 0
    gCancelRan = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async4")
    Set opValue = app.Async("opc").WithCancellation
    opValue.RunsProc("TestReDimAsync.RecordWork") _
        .OnCancel("TestReDimAsync.RecordCancel").TracksState "st"
    app.Button("btn").At("B2:C3").Text("Guarded")
    app.Async("opc").Disables "btn"
    app.Render

    app.Async("opc").Start
    transcript = "statusAfterStart=" & app.State("st")
    app.CancelAsync "opc"
    ReDimUI.PumpOnce
    transcript = transcript & "|statusAfterCancel=" & app.State("st")
    transcript = transcript & "|workNeverRan=" & CStr(gWorkRan = 0)
    transcript = transcript & "|cancelRan=" & gCancelRan
    transcript = transcript & "|buttonRestored=" & _
        CStr(Not app.Button("btn").IsBusy)

    ' A canceled cooperative delay: cancel mid-flight on a transport task.
    Set opValue = app.Async("opd").WithCancellation
    opValue.RunsTask(ROneCOne.Task.Delay(5000, opValue.Token)).TracksState "st2"
    app.Async("opd").Start
    ReDimUI.PumpOnce
    transcript = transcript & "|midFlight=" & app.State("st2")
    app.CancelAsync "opd"
    ReDimUI.PumpOnce
    transcript = transcript & "|midFlightCanceled=" & app.State("st2")
    ReDimUI.AutoPump True
    TestCancellation = transcript
End Function

Public Function TestJobChunks() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim firstTickCount As Long
    Dim safety As Long
    Dim transcript As String

    Set host = NewCanvas()
    gJobCounter = 0
    gDoneRan = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async5")
    app.Job("imp").Steps("TestReDimAsync.JobStepSmall").BudgetMs(5) _
        .JobOnDone "TestReDimAsync.RecordDone"
    app.Job("imp").StartJob

    ReDimUI.PumpOnce
    firstTickCount = gJobCounter
    transcript = "firstTickSteps=" & firstTickCount
    transcript = transcript & "|chunked=" & _
        CStr(firstTickCount >= 1 And firstTickCount < 5)
    safety = 0
    Do While app.Job("imp").JobIsRunning And safety < 20
        ReDimUI.PumpOnce
        safety = safety + 1
    Loop
    transcript = transcript & "|finalSteps=" & gJobCounter
    transcript = transcript & "|doneRan=" & gDoneRan
    transcript = transcript & "|ticksUsed=" & (safety + 1)
    ReDimUI.AutoPump True
    TestJobChunks = transcript
End Function

Public Function TestJobCancel() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gJobCounter = 0
    gCancelRan = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async6")
    app.Job("big").Steps("TestReDimAsync.JobStepEndless").BudgetMs(5) _
        .JobOnCancel "TestReDimAsync.RecordCancel"
    app.Job("big").StartJob

    ReDimUI.PumpOnce
    transcript = "partialSteps=" & gJobCounter
    app.CancelJob "big"
    ReDimUI.PumpOnce
    transcript = transcript & "|cancelRan=" & gCancelRan
    transcript = transcript & "|stoppedEarly=" & CStr(gJobCounter < 1000)
    transcript = transcript & "|stillRunning=" & _
        CStr(app.Job("big").JobIsRunning)
    ReDimUI.AutoPump True
    TestJobCancel = transcript
End Function

' The crown: a real armed SetTimer completes an op with no PumpOnce calls,
' then disarms itself once the work drains.
Public Function TestRealTimerEndToEnd() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim ignored As Variant
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "async7")
    app.Async("real").RunsTask(ROneCOne.Task.Delay(200)).TracksState "est"
    app.Render
    app.Async("real").Start

    transcript = "armedAfterStart=" & CStr(RdxPumpArmed())
    transcript = transcript & "|cursorPinned=" & _
        CStr(Application.Cursor = xlNorthwestArrow)
    ignored = ROneCOne.Task.Delay(700).Await
    transcript = transcript & "|statusNoManualPump=" & app.State("est")
    transcript = transcript & "|autoDisarmed=" & CStr(Not RdxPumpArmed())
    transcript = transcript & "|cursorRestored=" & _
        CStr(Application.Cursor = xlDefault)
    RdxStopPump
    TestRealTimerEndToEnd = transcript
End Function
