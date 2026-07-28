Attribute VB_Name = "TestReDimAsync"
Option Explicit

' Live scenarios for the ReDim async engine: ops, the OnClickAsync path,
' transport tasks across ticks, cancellation, chunked jobs, and one real
' SetTimer end-to-end run. Deterministic scenarios turn AutoPump off and
' drive PumpOnce; the end-to-end scenario uses the real timer and proves it
' disarms itself before returning.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long) As Long

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
    transcript = transcript & "|spinnerIsRing=" & _
        CStr(host.Shapes("rdm_async1_spn").AutoShapeType = msoShapeBlockArc)
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

' A default-budget job yields to the frame rate while something animates.
' The armed pump raises the timer resolution in production; PumpOnce does
' not arm a timer, so the scenario raises it the same way, or budget
' arithmetic below ~16 ms would quantize into a single step either way.
Public Function TestAdaptiveBudget() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim stepsIdle As Long
    Dim stepsAnimating As Long
    Dim transcript As String

    timeBeginPeriod 1
    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async8")
    app.Spinner("spn").AtRect 24, 24, 26, 26
    app.Render

    gJobCounter = 0
    app.Job("crunch").Steps("TestReDimAsync.JobStepEndless") _
        .JobOnCancel "TestReDimAsync.RecordCancel"
    app.Job("crunch").StartJob
    ReDimUI.PumpOnce
    stepsIdle = gJobCounter
    app.CancelJob "crunch"
    ReDimUI.PumpOnce

    app.Spinner("spn").Visible True
    gJobCounter = 0
    app.Job("crunch").StartJob
    ReDimUI.PumpOnce
    stepsAnimating = gJobCounter
    app.CancelJob "crunch"
    ReDimUI.PumpOnce
    app.Spinner("spn").Visible False

    transcript = "stepsIdle=" & stepsIdle
    transcript = transcript & "|stepsAnimating=" & stepsAnimating
    transcript = transcript & "|yieldsToAnimation=" & _
        CStr(stepsAnimating < stepsIdle)
    ReDimUI.AutoPump True
    timeEndPeriod 1
    TestAdaptiveBudget = transcript
End Function

' The spinner rotates every animation frame, and its frame must hold at
' the model position through sustained spinning: a walking spinner was
' reported in interactive use, so the animation tick re-pins the frame.
' The nudge at the end proves the pin repairs a deviation, whatever
' causes one, within a single frame.
Public Function TestSpinnerFramePinned() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim spinnerShape As Shape
    Dim rotBefore As Double
    Dim round As Long
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async10")
    app.Spinner("spn").AtRect 137.37, 42.61, 26, 26
    app.Spinner("spn").Visible True
    app.Render

    Set spinnerShape = host.Shapes("rdm_async10_spn")
    rotBefore = spinnerShape.Rotation
    For round = 1 To 40
        Sleep 20
        ReDimUI.PumpOnce
    Next round
    transcript = "spun=" & CStr(spinnerShape.Rotation <> rotBefore)
    transcript = transcript & "|leftPinned=" & _
        CStr(Abs(spinnerShape.Left - 137.37) <= 0.02)
    transcript = transcript & "|topPinned=" & _
        CStr(Abs(spinnerShape.Top - 42.61) <= 0.02)

    spinnerShape.Left = 300
    spinnerShape.Top = 200
    Sleep 20
    ReDimUI.PumpOnce
    transcript = transcript & "|repinned=" & _
        CStr(Abs(spinnerShape.Left - 137.37) <= 0.02 And _
            Abs(spinnerShape.Top - 42.61) <= 0.02)
    app.Spinner("spn").Visible False
    ReDimUI.AutoPump True
    TestSpinnerFramePinned = transcript
End Function

' Rendering a slider must arm the pump for its press watch: demand alone
' only keeps an already-armed pump alive, so a fresh workbook whose only
' interactive content is a slider would otherwise never watch for drags.
Public Function TestRenderArmsDragWatch() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "async9")
    app.SlideBar("vol").AtRect 24, 24, 200, 18
    app.Render
    transcript = "armedAfterRender=" & CStr(RdxPumpArmed())
    transcript = transcript & "|demandHolds=" & CStr(ReDimUI.HasPendingWork)
    app.Unmount True
    RdxStopPump
    transcript = transcript & "|cleanedUp=" & CStr(Not RdxPumpArmed())
    TestRenderArmsDragWatch = transcript
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
    ' Pinning is opt-in as of 0.7.0: unpinned by default so hover cursors
    ' survive, pinned on request, released again on demand.
    transcript = transcript & "|unpinnedByDefault=" & _
        CStr(Application.Cursor = xlDefault)
    ReDimUI.PinPumpCursor True
    transcript = transcript & "|pinsOnRequest=" & _
        CStr(Application.Cursor = xlNorthwestArrow)
    ReDimUI.PinPumpCursor False
    transcript = transcript & "|unpinsOnRequest=" & _
        CStr(Application.Cursor = xlDefault)
    ignored = ROneCOne.Task.Delay(700).Await
    transcript = transcript & "|statusNoManualPump=" & app.State("est")
    transcript = transcript & "|autoDisarmed=" & CStr(Not RdxPumpArmed())
    RdxStopPump
    TestRealTimerEndToEnd = transcript
End Function
