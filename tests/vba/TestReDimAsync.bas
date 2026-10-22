Attribute VB_Name = "TestReDimAsync"
Option Explicit

' Live scenarios for the ReDim async engine: ops, the OnClickAsync path,
' transport tasks across ticks, cancellation, chunked jobs, and real
' SetTimer runs: end to end, and held while code runs. Deterministic
' scenarios turn AutoPump off and drive PumpOnce; the real-timer scenarios
' prove the timer disarmed before returning. The Debugger procedures at
' the end stop in the debugger on purpose, and only test_pump_debugger.py
' runs them, each in an Excel of its own.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal windowHandle As LongPtr, ByVal ancestorKind As Long) As LongPtr
Private Declare PtrSafe Function GetClassNameW Lib "user32" ( _
    ByVal windowHandle As LongPtr, ByVal classBuffer As LongPtr, _
    ByVal bufferChars As Long) As Long

Private gWorkRan As Long
Private gDoneRan As Long
Private gCancelRan As Long
Private gDuringWorkBusy As Boolean
Private gJobCounter As Long
Private gResultSeen As Variant
Private gTagSeen As String
Private gStepTag As String
Private gPumpTicks As Long
Private gDebuggerTicks As Long
Private gStopInStep As Boolean

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

' An outcome handler that reads its op's result through the sender.
Public Sub ReadResult()
    gResultSeen = ReDimUI.Sender.Task.Result
End Sub

Public Sub RecordTag()
    gTagSeen = CStr(ReDimUI.Sender.TagValue)
End Sub

Public Function TaggedStep() As Boolean
    gStepTag = CStr(ReDimUI.Sender.TagValue) & "/" & ReDimUI.SenderId
    TaggedStep = True
End Function

' A paced step that unmounts its own app and reports itself finished.
Public Function UnmountingStep() As Boolean
    ReDimUI.App("async12").Unmount True
    UnmountingStep = True
End Function

' A Tag reaches the handlers through ReDimUI.Sender: a button's click, a
' job's step, and the job's outcome handler.
Public Function TestTagsReachHandlers() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gTagSeen = vbNullString
    gStepTag = vbNullString
    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async11")
    app.Button("go").At("B2:C3").Text("Go").OnClick("TestReDimAsync.RecordTag").Tag "fromButton"
    app.Job("tagged").Tag(7).Steps("TestReDimAsync.TaggedStep").JobOnDone "TestReDimAsync.RecordTag"
    app.Render
    ReDimUI.DispatchShape "rdm_async11_go"
    transcript = "buttonTag=" & gTagSeen
    app.Job("tagged").StartJob
    ReDimUI.PumpOnce
    transcript = transcript & "|stepTag=" & gStepTag & "|doneTag=" & gTagSeen

    ' A paced step that unmounts its app stops its job: OnDone does not
    ' run against the app that is gone.
    gDoneRan = 0
    Set app = ReDimUI.Mount(NewCanvas(), "async12")
    app.Job("selfEnd").Steps("TestReDimAsync.UnmountingStep").PacedMs(1) _
        .JobOnDone "TestReDimAsync.RecordDone"
    app.Render
    app.Job("selfEnd").StartJob
    Sleep 5
    ReDimUI.PumpOnce
    transcript = transcript & "|pacedUnmount=" & gDoneRan & "/" & _
        CStr(Not ReDimUI.HasApp("async12"))
    ReDimUI.AutoPump True
    TestTagsReachHandlers = transcript
End Function

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

    ' The body runs in the pump with the button clicked as Sender, so one
    ' procedure serves buttons told apart by their Tag.
    gTagSeen = vbNullString
    app.Button("tagged").AtRect(200, 24, 90, 28).Text("B").Tag("B") _
        .OnClickAsync "TestReDimAsync.RecordTag"
    ReDimUI.DispatchShape "rdm_async2_tagged"
    ReDimUI.PumpOnce
    transcript = transcript & "|bodySender=" & gTagSeen

    ' A click op at rest has nothing to cancel, token source or none.
    On Error Resume Next
    app.CancelAsync "click_btnrun"
    transcript = transcript & "|restCancelQuiet=" & CStr(Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    ReDimUI.AutoPump True
    TestOnClickAsyncSugar = transcript
End Function

' An op's outcome handler runs with the op as ReDimUI.Sender and reads
' its task's result through Sender.Task. (An error raised inside a
' handler escapes a harness call whatever traps it, so the error sink is
' not exercised here.)
Public Function TestOutcomeHandlers() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gResultSeen = Empty
    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async9")
    app.Async("val").RunsTask(ROneCOne.Task.FromResult(42)).OnDone "TestReDimAsync.ReadResult"
    app.Render
    app.Async("val").Start
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = "senderTask=" & CStr(gResultSeen)
    ReDimUI.AutoPump True
    TestOutcomeHandlers = transcript
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

' A cancelled op runs again: its next start, and the Token read for it,
' get a fresh token, and cancelling an op at rest, or an id never made,
' changes nothing. IsRunning reads the op's run. An app that unmounts
' stops its ops and jobs without their handlers.
Public Function TestRunsAgainAfterCancel() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gWorkRan = 0
    gCancelRan = 0
    gDoneRan = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async10")
    app.Async("opr").WithCancellation.RunsProc("TestReDimAsync.RecordWork") _
        .OnCancel("TestReDimAsync.RecordCancel").TracksState "st"
    app.Render
    app.Async("opr").Start
    transcript = "runningRead=" & CStr(app.Async("opr").IsRunning)
    app.CancelAsync "opr"
    app.CancelAsync "neverMade"
    app.CancelJob "neverMade"
    ReDimUI.PumpOnce
    transcript = transcript & "|firstCanceled=" & app.State("st") & "/" & gCancelRan & "/" & _
        CStr(app.Async("opr").IsRunning)
    transcript = transcript & "|freshToken=" & _
        CStr(Not app.Async("opr").Token.IsCancellationRequested)
    app.Async("opr").Start
    ReDimUI.PumpOnce
    transcript = transcript & "|runsAgain=" & app.State("st") & "/" & gWorkRan
    app.CancelAsync "opr"
    app.Async("opr").Start
    ReDimUI.PumpOnce
    transcript = transcript & "|idleCancelIgnored=" & app.State("st") & "/" & gWorkRan
    gJobCounter = 0
    app.Job("endless").Steps("TestReDimAsync.JobStepEndless") _
        .JobOnCancel "TestReDimAsync.RecordCancel"
    app.Job("endless").StartJob
    app.Unmount True
    transcript = transcript & "|unmountStops=" & CStr(Not app.Job("endless").IsRunning) & _
        "/" & gCancelRan
    ReDimUI.AutoPump True
    TestRunsAgainAfterCancel = transcript
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
    Do While app.Job("imp").IsRunning And safety < 20
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
        CStr(app.Job("big").IsRunning)
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
    ROneCOne.Task.Delay(700).Await
    transcript = transcript & "|statusNoManualPump=" & app.State("est")
    transcript = transcript & "|autoDisarmed=" & CStr(Not RdxPumpArmed())
    RdxStopPump
    TestRealTimerEndToEnd = transcript
End Function

' Ops that share a spinner and a control they disable keep both busy
' until the last of them finishes, and a dialog waiting on its own asks
' nothing of the pump.
Public Function TestSharedBusy() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "async12")
    app.Button("next").AtRect(24, 24, 90, 28).Text "Next"
    app.Spinner("spn").AtRect(130, 24, 24, 24).Visible False
    app.Async("short").RunsTask(ROneCOne.Task.Delay(50)).ShowsSpinner("spn").Disables "next"
    app.Async("long").RunsTask(ROneCOne.Task.Delay(600)).ShowsSpinner("spn").Disables "next"
    app.Render
    app.Async("short").Start
    app.Async("long").Start
    Sleep 150
    ReDimUI.PumpOnce
    transcript = "whileOneRuns=" & CStr(app.Async("short").OpIsRunning) & "/" & _
        CStr(app.Button("next").IsBusy) & "/" & CStr(app.Spinner("spn").IsVisible)
    Sleep 600
    ReDimUI.PumpOnce
    transcript = transcript & "|afterLast=" & CStr(app.Button("next").IsBusy) & "/" & _
        CStr(app.Spinner("spn").IsVisible)

    app.Confirm "Proceed?", "Go on?", "TestReDimAsync.RecordDone"
    transcript = transcript & "|dialogAsksPump=" & CStr(ReDimUI.HasPendingWork)
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    RdxStopPump
    ReDimUI.AutoPump True
    TestSharedBusy = transcript
End Function

' A paced step that counts the passes the real timer drives, and never
' finishes.
Public Function CountingStep() As Boolean
    gPumpTicks = gPumpTicks + 1
    CountingStep = False
End Function

' Code that holds the pump keeps its timer from firing: armed before a
' hold or under one, the timer runs no tick until every hold lets go.
' Each entry point, and each handler ReDim runs, lets go of its hold.
Public Function TestPumpHolds() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim ticksBefore As Long

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "async13")
    app.Button("tap").AtRect(24, 24, 90, 28).Text("Tap").OnClick "TestReDimAsync.RecordWork"
    app.OnStateChanged "seen", "TestReDimAsync.RecordDone"
    app.Render
    app.Job("count").Steps("TestReDimAsync.CountingStep").PacedMs 1
    gPumpTicks = 0
    app.Job("count").StartJob
    ROneCOne.Task.Delay(300).Await
    transcript = "ticks=" & CStr(gPumpTicks > 0)

    RdxHoldPump
    ticksBefore = gPumpTicks
    ROneCOne.Task.Delay(300).Await
    transcript = transcript & "|heldStill=" & CStr(gPumpTicks = ticksBefore)
    RdxHoldPump
    RdxReleasePump
    ROneCOne.Task.Delay(300).Await
    transcript = transcript & "|nestedHeld=" & CStr(gPumpTicks = ticksBefore)
    RdxReleasePump
    ROneCOne.Task.Delay(300).Await
    transcript = transcript & "|released=" & CStr(gPumpTicks > ticksBefore)

    RdxStopPump
    RdxHoldPump
    RdxEnsurePump
    ticksBefore = gPumpTicks
    ROneCOne.Task.Delay(300).Await
    transcript = transcript & "|armedUnderHold=" & CStr(RdxPumpArmed()) & "/" & _
        CStr(gPumpTicks = ticksBefore)
    RdxReleasePump
    ROneCOne.Task.Delay(300).Await
    transcript = transcript & "|ticksAfterHold=" & CStr(gPumpTicks > ticksBefore)

    RdxKeyChar "{ESC}"
    RdxShortcut "^+q"
    RdxAccessKey "q"
    RdxOpenPalette
    RdxPalettePick
    RdxDispatch
    transcript = transcript & "|entriesLetGo=" & CStr(Not RdxPumpHeld())
    ReDimUI.DispatchShape "rdm_async13_tap"
    app.SetState "seen", 1
    transcript = transcript & "|handlersLetGo=" & CStr(Not RdxPumpHeld())

    Do While RdxPumpHeld()
        RdxReleasePump
    Loop
    app.CancelJob "count"
    app.Unmount True
    RdxStopPump
    TestPumpHolds = transcript
End Function

' The pump waits while Excel's active window is not a workbook's own, such
' as the VBA editor, and comes back with the next click, key, or selection
' once a workbook window is active again. The seam stands in for Excel's
' active window: a background Excel cannot hand activation back itself.
Public Function TestPumpPauses() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim ticksBefore As Long
    Dim editorWindow As LongPtr

    editorWindow = Application.VBE.MainWindow.HWnd
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "async14")
    app.Job("count").Steps("TestReDimAsync.CountingStep").PacedMs 1
    gPumpTicks = 0
    app.Job("count").StartJob
    ROneCOne.Task.Delay(200).Await
    transcript = "ticks=" & CStr(gPumpTicks > 0)

    RdxOverrideActiveWindow editorWindow
    ROneCOne.Task.Delay(100).Await
    transcript = transcript & "|pausedForEditor=" & CStr(RdxPumpPaused()) & "/" & _
        CStr(Not RdxPumpArmed())
    ticksBefore = gPumpTicks
    ROneCOne.Task.Delay(200).Await
    transcript = transcript & "|waits=" & CStr(gPumpTicks = ticksBefore)
    RdxKeyChar "{ESC}"
    RdxEnsurePump
    transcript = transcript & "|keyWhileEditor=" & CStr(RdxPumpPaused()) & "/" & _
        CStr(Not RdxPumpArmed())

    RdxOverrideActiveWindow Application.Hwnd
    ROneCOne.Task.Delay(200).Await
    transcript = transcript & "|stillWaits=" & CStr(RdxPumpPaused())
    RdxKeyChar "{ESC}"
    transcript = transcript & "|keyResumes=" & CStr(RdxPumpArmed()) & "/" & _
        CStr(Not RdxPumpPaused())
    ticksBefore = gPumpTicks
    ROneCOne.Task.Delay(200).Await
    transcript = transcript & "|ticksAgain=" & CStr(gPumpTicks > ticksBefore)

    ' Stepping out of a handler in the debugger lets go of its hold with
    ' the editor active.
    RdxHoldPump
    RdxOverrideActiveWindow editorWindow
    RdxReleasePump
    transcript = transcript & "|releaseIntoEditor=" & CStr(RdxPumpPaused()) & "/" & _
        CStr(Not RdxPumpArmed())
    ' No window active, as in a hidden Excel, leaves the pump running.
    RdxOverrideActiveWindow 0
    RdxDispatch
    transcript = transcript & "|noWindowRuns=" & CStr(RdxPumpArmed())

    RdxClearActiveWindowOverride
    app.CancelJob "count"
    app.Unmount True
    RdxStopPump
    TestPumpPauses = transcript
End Function

' Live debugger scenarios. A paced job keeps the real timer armed, as a
' live app's pointer watch or focused field does, before a stop in the
' debugger or an edit in the VBA editor.
Public Function DebuggerSetup() As String
    Dim app As ReDimUI

    gDebuggerTicks = 0
    gStopInStep = False
    ReDimUI.AutoPump True
    Set app = ReDimUI.Mount(NewCanvas(), "debugger")
    app.Button("halt").AtRect(24, 24, 90, 28).Text("Halt").OnClick "TestReDimAsync.StopInHandler"
    app.Render
    app.Job("beat").Steps("TestReDimAsync.DebuggerStep").PacedMs 1
    app.Job("beat").StartJob
    DebuggerSetup = DebuggerState()
End Function

Public Function DebuggerState() As String
    DebuggerState = "armed=" & CStr(RdxPumpArmed()) & "|paused=" & _
        CStr(RdxPumpPaused()) & "|ticks=" & CStr(gDebuggerTicks) & _
        "|active=" & ActiveWindowClass()
End Function

' The class of the window that owns Excel's active window, as the pump
' reads it; empty with no window active.
Private Function ActiveWindowClass() As String
    Dim activeHandle As LongPtr
    Dim classBuffer As String
    Dim classLength As Long

    activeHandle = GetActiveWindow()
    If activeHandle = 0 Then Exit Function
    classBuffer = String$(64, vbNullChar)
    classLength = GetClassNameW(GetAncestor(activeHandle, 3), StrPtr(classBuffer), 64)
    ActiveWindowClass = Left$(classBuffer, classLength)
End Function

Public Function DebuggerStep() As Boolean
    gDebuggerTicks = gDebuggerTicks + 1
    If gStopInStep Then
        gStopInStep = False
        Stop
    End If
    DebuggerStep = False
End Function

Public Sub StopInHandler()
    Stop
End Sub

' The click comes once this call has returned, as a user's would.
Public Sub DebuggerClickLater()
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!TestReDimAsync.DebuggerClick"
End Sub

Public Sub DebuggerClick()
    ReDimUI.DispatchShape "rdm_debugger_halt"
End Sub

Public Sub DebuggerStopInStep()
    gStopInStep = True
End Sub

' The VBA editor comes forward on ReDimHost with the caret at the start of
' the TIMERPROC's body, as for a user about to type there. Returns the
' editor's window.
Public Function DebuggerOpenHostCode() As String
    Dim editor As Object
    Dim hostCode As Object
    Dim bodyLine As Long

    Set editor = Application.VBE
    editor.MainWindow.Visible = True
    editor.MainWindow.SetFocus
    Set hostCode = ThisWorkbook.VBProject.VBComponents("ReDimHost").CodeModule
    hostCode.CodePane.Show
    bodyLine = hostCode.ProcBodyLine("RdxPumpCallback", 0)
    Do While Right$(RTrim$(hostCode.Lines(bodyLine, 1)), 1) = "_"
        bodyLine = bodyLine + 1
    Loop
    hostCode.CodePane.SetSelection bodyLine + 1, 1, bodyLine + 1, 1
    DebuggerOpenHostCode = CStr(editor.MainWindow.HWnd)
End Function

' Excel shows, so a workbook window is its active window.
Public Sub DebuggerShowExcel()
    Application.Visible = True
End Sub

Public Sub DebuggerFocusEditor()
    Application.VBE.MainWindow.Visible = True
    Application.VBE.MainWindow.SetFocus
End Sub

' A selection the user makes, with events on as in a live session.
Public Sub DebuggerSelectCell()
    Dim eventsWere As Boolean

    eventsWere = Application.EnableEvents
    Application.EnableEvents = True
    ActiveSheet.Range("H20").Select
    Application.EnableEvents = eventsWere
End Sub

Public Function DebuggerTeardown() As String
    ReDimUI.App("debugger").CancelJob "beat"
    ReDimUI.App("debugger").Unmount True
    RdxStopPump
    Application.VBE.MainWindow.Visible = False
    Application.Visible = False
    DebuggerTeardown = DebuggerState()
End Function
