Attribute VB_Name = "TestReDimWidgets"
Option Explicit

' Live scenarios for the widget set: progress, toggle, form controls, cell
' inputs, toasts, and the shapes-based modal.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private gChangeCount As Long
Private gConfirmRan As Long
Private gCancelRan As Long

Private Function NewCanvas() As Worksheet
    Set NewCanvas = ActiveWorkbook.Worksheets.Add
End Function

Public Sub RecordChange()
    gChangeCount = gChangeCount + 1
End Sub

Public Sub RecordConfirm()
    gConfirmRan = gConfirmRan + 1
End Sub

Public Sub RecordCancelChoice()
    gCancelRan = gCancelRan + 1
End Sub

Public Function TestProgressBar() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim track As Shape
    Dim fillPart As Shape
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid1")
    app.ProgressBar("prg").At("B2:F2").BindValue "pct"
    app.SetState "pct", 0
    app.Render

    Set track = host.Shapes("rdm_wid1_prg")
    transcript = "zeroHidesFill=" & _
        CStr(host.Shapes("rdm_wid1_prg__fill").Visible = msoFalse)
    app.SetState "pct", 50
    Set fillPart = host.Shapes("rdm_wid1_prg__fill")
    transcript = transcript & "|halfVisible=" & CStr(fillPart.Visible = msoTrue)
    transcript = transcript & "|halfWidthOk=" & _
        CStr(Abs(fillPart.Width - track.Width / 2) < 0.5)
    app.SetState "pct", 100
    transcript = transcript & "|fullWidthOk=" & _
        CStr(Abs(fillPart.Width - track.Width) < 0.5)
    app.SetState "pct", 250
    transcript = transcript & "|clampedOk=" & _
        CStr(Abs(fillPart.Width - track.Width) < 0.5)
    TestProgressBar = transcript
End Function

Public Function TestToggle() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim knob As Shape
    Dim transcript As String
    Dim leftBefore As Double

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid2")
    app.Toggle("tgl").At("B2").WritesTo("darkMode").OnChange _
        "TestReDimWidgets.RecordChange"
    app.SetState "darkMode", False
    app.Render

    Set knob = host.Shapes("rdm_wid2_tgl__knob")
    leftBefore = knob.Left
    transcript = "offFillMuted=" & _
        CStr(host.Shapes("rdm_wid2_tgl").Fill.ForeColor.RGB = _
            app.Theme.MutedColor)
    ReDimUI.DispatchShape "rdm_wid2_tgl"
    transcript = transcript & "|stateOn=" & CStr(app.State("darkMode"))
    transcript = transcript & "|changeRan=" & gChangeCount
    transcript = transcript & "|onFillPrimary=" & _
        CStr(host.Shapes("rdm_wid2_tgl").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)
    transcript = transcript & "|knobMoved=" & CStr(knob.Left > leftBefore)
    transcript = transcript & "|checkedProp=" & _
        CStr(app.Toggle("tgl").IsChecked)
    ' Clicking the knob part resolves to the same component.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid2_tgl__knob"
    transcript = transcript & "|knobClickTogglesOff=" & _
        CStr(app.State("darkMode") = False)
    TestToggle = transcript
End Function

Public Function TestSelectBox() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim faceShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid7")
    app.SelectBox("pick").AtRect(24, 24, 130, 24) _
        .Items("North", "South", "East", "West").Value(2).WritesTo "region"
    app.SelectBox("pick").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    Set faceShape = host.Shapes("rdm_wid7_pick")
    transcript = "faceText=" & faceShape.TextFrame2.TextRange.Text
    transcript = transcript & "|faceFontSize=" & _
        faceShape.TextFrame2.TextRange.Font.Size
    transcript = transcript & "|faceInkOnSurface=" & _
        CStr(faceShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            app.Theme.OnSurfaceColor)
    transcript = transcript & "|caretExists=" & _
        CStr(ShapeExists(host, "rdm_wid7_pick__caret"))
    transcript = transcript & "|closedNoOptions=" & _
        CStr(Not ShapeExists(host, "rdm_wid7_pick__opt1"))

    ReDimUI.DispatchShape "rdm_wid7_pick"
    transcript = transcript & "|openOptions=" & _
        CStr(ShapeExists(host, "rdm_wid7_pick__opt1") And _
             ShapeExists(host, "rdm_wid7_pick__opt4"))
    transcript = transcript & "|optionFontSize=" & _
        host.Shapes("rdm_wid7_pick__opt3").TextFrame2.TextRange.Font.Size
    transcript = transcript & "|optionText=" & _
        host.Shapes("rdm_wid7_pick__opt3").TextFrame2.TextRange.Text

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick__opt3"
    transcript = transcript & "|pickedState=" & app.State("region")
    transcript = transcript & "|pickedFace=" & _
        faceShape.TextFrame2.TextRange.Text
    transcript = transcript & "|closedAfterPick=" & _
        CStr(Not ShapeExists(host, "rdm_wid7_pick__opt1"))
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Reopen and close by clicking the face again.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick"
    transcript = transcript & "|toggleClosed=" & _
        CStr(Not ShapeExists(host, "rdm_wid7_pick__opt1"))
    TestSelectBox = transcript
End Function

Public Function TestTextInput() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid4")
    app.TextInput("name").At("C3").WritesTo("userName").OnChange _
        "TestReDimWidgets.RecordChange"
    app.Render

    transcript = "frameExists=" & _
        CStr(Not host.Shapes("rdm_wid4_name") Is Nothing)
    app.TextInput("name").InputValue = "Ada"
    transcript = transcript & "|apiWriteState=" & app.State("userName")
    transcript = transcript & "|apiNoChangeProc=" & CStr(gChangeCount = 0)

    ' A real cell edit fires Application.SheetChange into the framework. The
    ' harness may hold EnableEvents off, so the scenario pins it on and
    ' restores it after, recording what it found for the transcript.
    Dim eventsWereOn As Boolean
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = True
    host.Range("C3").Value = "Grace"
    Application.EnableEvents = eventsWereOn
    transcript = transcript & "|eventsWereOn=" & CStr(eventsWereOn)
    transcript = transcript & "|editState=" & app.State("userName")
    transcript = transcript & "|editChangeProc=" & CStr(gChangeCount = 1)
    transcript = transcript & "|inputReadback=" & app.TextInput("name").InputValue
    TestTextInput = transcript
End Function

Public Function TestToastLifecycle() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim toastValue As ReDimUI
    Dim toastName As String
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid5")
    app.Label("anchorlbl").At("B2").Text("x")
    app.Render

    Set toastValue = app.Toast("Saved.", 150)
    toastName = "rdm_wid5_" & toastValue.ComponentId
    transcript = "shown=" & CStr(host.Shapes(toastName).Visible = msoTrue)
    transcript = transcript & "|pendingWork=" & CStr(ReDimUI.HasPendingWork)
    ReDimUI.PumpOnce
    transcript = transcript & "|aliveBeforeTtl=" & CStr(ShapeExists(host, toastName))
    Sleep 200
    ReDimUI.PumpOnce
    transcript = transcript & "|removedAfterTtl=" & _
        CStr(Not ShapeExists(host, toastName))

    ' Click dismissal: a fresh toast dies on the tick after its click.
    Set toastValue = app.Toast("Click me.", 60000)
    toastName = "rdm_wid5_" & toastValue.ComponentId
    ReDimUI.DispatchShape toastName
    ReDimUI.PumpOnce
    transcript = transcript & "|clickDismissed=" & _
        CStr(Not ShapeExists(host, toastName))

    ' Toast ink must stay readable in both themes.
    Set toastValue = app.Toast("Readable.", 60000)
    toastName = "rdm_wid5_" & toastValue.ComponentId
    transcript = transcript & "|lightInk=" & _
        CStr(host.Shapes(toastName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB _
            = app.Theme.OnSurfaceColor)
    app.SetTheme ReDimUI.ThemeDark
    transcript = transcript & "|darkInk=" & _
        CStr(host.Shapes(toastName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB _
            = app.Theme.OnSurfaceColor)
    ReDimUI.AutoPump True
    TestToastLifecycle = transcript
End Function

Private Function ShapeExists(ByVal host As Worksheet, ByVal shapeName As String) As Boolean
    Dim probe As Shape

    On Error Resume Next
    Set probe = host.Shapes(shapeName)
    On Error GoTo 0
    ShapeExists = Not probe Is Nothing
End Function

Public Function TestToastSlots() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim firstToast As ReDimUI
    Dim secondToast As ReDimUI
    Dim thirdToast As ReDimUI
    Dim firstName As String
    Dim firstSettledTop As Double
    Dim secondSettledTop As Double
    Dim ticks As Long
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid8")
    app.Label("anchor").AtRect(24, 24, 200, 20).Text("x")
    app.Render

    ' Entrance slide: a toast spawns 14 points low and eases into its slot.
    Set firstToast = app.Toast("one", 60000)
    firstName = "rdm_wid8_" & firstToast.ComponentId
    Dim entryTop As Double
    entryTop = host.Shapes(firstName).Top
    For ticks = 1 To 10
        ReDimUI.PumpOnce
    Next ticks
    transcript = "entranceSlid=" & _
        CStr(Abs((entryTop - host.Shapes(firstName).Top) - 14) < 0.1)

    Set secondToast = app.Toast("two", 60000)
    For ticks = 1 To 10
        ReDimUI.PumpOnce
    Next ticks
    firstSettledTop = host.Shapes(firstName).Top
    secondSettledTop = host.Shapes("rdm_wid8_" & secondToast.ComponentId).Top
    transcript = transcript & "|secondBelowFirst=" & _
        CStr(secondSettledTop > firstSettledTop)

    ' Dismissing the first toast compacts the stack: the survivor slides up
    ' into slot one.
    ReDimUI.DispatchShape firstName
    For ticks = 1 To 12
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|survivorSlidUp=" & _
        CStr(Abs(host.Shapes("rdm_wid8_" & secondToast.ComponentId).Top - _
            firstSettledTop) < 0.1)

    ' A new toast joins below the compacted stack, in slot two.
    Set thirdToast = app.Toast("three", 60000)
    For ticks = 1 To 10
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|thirdJoinsBelow=" & _
        CStr(Abs(host.Shapes("rdm_wid8_" & thirdToast.ComponentId).Top - _
            secondSettledTop) < 0.1)
    ReDimUI.AutoPump True
    TestToastSlots = transcript
End Function

' A dragged modal piece (or any component) must snap back to its declared
' rectangle, and framework shapes must swallow plain clicks so accidental
' drags cannot happen in the first place.
Public Function TestDragResilience() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim cardShape As Shape
    Dim declaredLeft As Double
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid9")
    app.Label("info").AtRect(24, 24, 200, 20).Text("Info")
    app.Render
    app.Confirm "Move me", "Try to drag this.", vbNullString

    Set cardShape = host.Shapes("rdm_wid9_mdl_card")
    declaredLeft = cardShape.Left
    transcript = "cardSwallowsClicks=" & _
        CStr(InStr(cardShape.OnAction, "RdxDispatch") > 0)
    transcript = transcript & "|labelSwallowsClicks=" & _
        CStr(InStr(host.Shapes("rdm_wid9_info").OnAction, "RdxDispatch") > 0)

    ' Simulate a manual drag, then reopen: geometry must snap back.
    cardShape.Left = declaredLeft + 140
    cardShape.Top = cardShape.Top + 60
    app.Confirm "Move me", "Try to drag this.", vbNullString
    transcript = transcript & "|cardSnappedBack=" & _
        CStr(Abs(host.Shapes("rdm_wid9_mdl_card").Left - declaredLeft) < 0.5)

    ' The same snap applies to ordinary components on Render.
    host.Shapes("rdm_wid9_info").Left = 300
    app.Render
    transcript = transcript & "|labelSnappedBack=" & _
        CStr(Abs(host.Shapes("rdm_wid9_info").Left - 24) < 0.01)
    app.CloseModal
    TestDragResilience = transcript
End Function

' The toast tray must be stable: pinned by ToastTray when given, and never
' shifted by modal chrome even when a toast fires while a modal is open.
Public Function TestToastTray() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim toastValue As ReDimUI
    Dim labelShape As Shape
    Dim expectedRail As Double
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid10")
    app.Label("info").AtRect(24, 24, 200, 20).Text("Info")
    app.Render

    ' Default rail: just outside the content's right edge.
    Set labelShape = host.Shapes("rdm_wid10_info")
    expectedRail = labelShape.Left + labelShape.Width + 12
    Set toastValue = app.Toast("rail", 60000)
    transcript = "onRail=" & _
        CStr(Abs(host.Shapes("rdm_wid10_" & toastValue.ComponentId).Left - _
            expectedRail) < 0.5)

    ' A modal must not shift the rail even while it is open.
    app.Confirm "Wait", "Working...", vbNullString
    Set toastValue = app.Toast("during modal", 60000)
    transcript = transcript & "|modalIgnored=" & _
        CStr(Abs(host.Shapes("rdm_wid10_" & toastValue.ComponentId).Left - _
            expectedRail) < 0.5)
    app.CloseModal

    ' Explicit tray anchor wins.
    app.ToastTray "H2"
    Set toastValue = app.Toast("pinned", 60000)
    transcript = transcript & "|pinnedToAnchor=" & _
        CStr(Abs(host.Shapes("rdm_wid10_" & toastValue.ComponentId).Left - _
            host.Range("H2").Left) < 0.5)

    ' Content wider than the window: the rail clamps into the viewport so
    ' the notification stays on screen.
    Dim wideApp As ReDimUI
    Dim wideHost As Worksheet
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Set wideHost = NewCanvas()
    Set wideApp = ReDimUI.Mount(wideHost, "wid11")
    wideApp.Label("far").AtRect(2000, 20, 300, 40).Text("Far away")
    wideApp.Render
    wideApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    Set toastValue = wideApp.Toast("clamped", 60000)
    Dim clampedShape As Shape
    Set clampedShape = wideHost.Shapes("rdm_wid11_" & toastValue.ComponentId)
    transcript = transcript & "|clampedIntoView=" & _
        CStr(clampedShape.Left + clampedShape.Width <= viewLeft + viewWidth _
            And clampedShape.Left >= viewLeft)
    transcript = transcript & "|notAtRawRail=" & _
        CStr(clampedShape.Left < 2312)
    ReDimUI.AutoPump True
    TestToastTray = transcript
End Function

Public Function TestTickBox() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim boxShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid12")
    app.TickBox("agree").AtRect(24, 24, 140, 18).Text("I agree") _
        .WritesTo("agreed").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    Set boxShape = host.Shapes("rdm_wid12_agree")
    transcript = "uncheckedSurface=" & _
        CStr(boxShape.Fill.ForeColor.RGB = app.Theme.SurfaceColor)
    transcript = transcript & "|glyphEmpty=" & _
        CStr(LenB(boxShape.TextFrame2.TextRange.Text) = 0)
    transcript = transcript & "|captionText=" & _
        host.Shapes("rdm_wid12_agree__lbl").TextFrame2.TextRange.Text
    transcript = transcript & "|captionSize=" & _
        host.Shapes("rdm_wid12_agree__lbl").TextFrame2.TextRange.Font.Size

    ReDimUI.DispatchShape "rdm_wid12_agree"
    transcript = transcript & "|checkedState=" & CStr(app.State("agreed"))
    transcript = transcript & "|checkedPrimary=" & _
        CStr(boxShape.Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    transcript = transcript & "|glyphCheck=" & _
        CStr(boxShape.TextFrame2.TextRange.Text = ChrW(10003))
    transcript = transcript & "|changeRan=" & gChangeCount

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid12_agree__lbl"
    transcript = transcript & "|captionToggles=" & _
        CStr(app.State("agreed") = False)
    TestTickBox = transcript
End Function

Public Function TestRadioGroup() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid13")
    app.RadioGroup("prio").AtRect(24, 24, 140, 60) _
        .Items("Low", "Medium", "High").Value(2).WritesTo("prio") _
        .OnChange "TestReDimWidgets.RecordChange"
    app.Render

    transcript = "rowsExist=" & _
        CStr(ShapeExists(host, "rdm_wid13_prio") And _
             ShapeExists(host, "rdm_wid13_prio__c2") And _
             ShapeExists(host, "rdm_wid13_prio__c3"))
    transcript = transcript & "|dotOnSelected=" & _
        CStr(host.Shapes("rdm_wid13_prio__d2").Visible = msoTrue)
    transcript = transcript & "|dotOffOthers=" & _
        CStr(host.Shapes("rdm_wid13_prio__d1").Visible = msoFalse And _
             host.Shapes("rdm_wid13_prio__d3").Visible = msoFalse)
    transcript = transcript & "|captionText=" & _
        host.Shapes("rdm_wid13_prio__t3").TextFrame2.TextRange.Text
    transcript = transcript & "|captionSize=" & _
        host.Shapes("rdm_wid13_prio__t3").TextFrame2.TextRange.Font.Size

    ReDimUI.DispatchShape "rdm_wid13_prio__t3"
    transcript = transcript & "|pickedState=" & app.State("prio")
    transcript = transcript & "|dotMoved=" & _
        CStr(host.Shapes("rdm_wid13_prio__d3").Visible = msoTrue And _
             host.Shapes("rdm_wid13_prio__d2").Visible = msoFalse)
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Clicking the already-selected row is a no-op.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid13_prio__c3"
    transcript = transcript & "|sameRowNoOp=" & CStr(gChangeCount = 1)

    ' Row one is reachable through the main circle shape.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid13_prio"
    transcript = transcript & "|rowOnePicked=" & app.State("prio")
    TestRadioGroup = transcript
End Function

Public Function TestStepper() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim faceShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid14")
    app.Stepper("thr").AtRect(24, 24, 120, 24).SliderRange(1, 5, 1) _
        .Value(4).WritesTo("threads").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    Set faceShape = host.Shapes("rdm_wid14_thr")
    transcript = "faceValue=" & faceShape.TextFrame2.TextRange.Text
    transcript = transcript & "|partsExist=" & _
        CStr(ShapeExists(host, "rdm_wid14_thr__minus") And _
             ShapeExists(host, "rdm_wid14_thr__plus"))

    ReDimUI.DispatchShape "rdm_wid14_thr__plus"
    transcript = transcript & "|plusValue=" & app.State("threads")
    transcript = transcript & "|faceUpdated=" & _
        CStr(faceShape.TextFrame2.TextRange.Text = "5")

    ' Already at the maximum: another plus is a clamped no-op.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid14_thr__plus"
    transcript = transcript & "|clampedNoOp=" & CStr(gChangeCount = 1)

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid14_thr__minus"
    transcript = transcript & "|minusValue=" & app.State("threads")
    transcript = transcript & "|changeRan=" & gChangeCount
    TestStepper = transcript
End Function

Public Function TestSlideBar() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim trackShape As Shape
    Dim fillShape As Shape
    Dim thumbShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid15")
    app.SlideBar("vol").AtRect(24, 24, 200, 18).SliderRange(0, 100, 5) _
        .Value(40).WritesTo("volume").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    Set trackShape = host.Shapes("rdm_wid15_vol")
    Set fillShape = host.Shapes("rdm_wid15_vol__fill")
    Set thumbShape = host.Shapes("rdm_wid15_vol__thumb")
    transcript = "partsExist=" & _
        CStr(Not fillShape Is Nothing And Not thumbShape Is Nothing)
    transcript = transcript & "|fillFraction=" & _
        CStr(Abs(fillShape.Width - 80) < 0.5)
    transcript = transcript & "|thumbCentered=" & _
        CStr(Abs((thumbShape.Left + thumbShape.Width / 2) - (24 + 80)) < 0.5)

    ' The deterministic seam the drag loop drives: set from track fractions.
    app.Component("vol").SlideToFraction app, 0.75, True
    transcript = transcript & "|threeQuarterValue=" & app.State("volume")
    transcript = transcript & "|changeRan=" & gChangeCount
    transcript = transcript & "|fillMoved=" & _
        CStr(Abs(fillShape.Width - 150) < 0.5)

    ' Snapping: 0.52 across 0..100 step 5 lands on 50.
    app.Component("vol").SlideToFraction app, 0.52, False
    transcript = transcript & "|snappedValue=" & app.State("volume")
    transcript = transcript & "|noExtraChange=" & CStr(gChangeCount = 1)

    ' Clamping at the rails.
    app.Component("vol").SlideToFraction app, 0#, False
    transcript = transcript & "|minValue=" & app.State("volume")
    transcript = transcript & "|fillHiddenAtMin=" & _
        CStr(fillShape.Visible = msoFalse)
    app.Component("vol").SlideToFraction app, 1#, False
    transcript = transcript & "|maxValue=" & app.State("volume")
    ' Render armed the pump for the slider's press watch; never leave a
    ' timer armed across harness call boundaries.
    RdxStopPump
    TestSlideBar = transcript
End Function

' The drag-session state machine through its deterministic seams: the
' press watch itself needs a physical mouse, but begin, live continue,
' release semantics, click swallowing, and pump demand are assertable.
Public Function TestSlideDrag() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim probe As ReDimUI
    Dim thumbShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid17")
    app.SlideBar("trk").AtRect 24, 24, 200, 18
    app.SlideBar("trk").SliderRange(0, 100, 5).Value(20).WritesTo("level") _
        .OnChange "TestReDimWidgets.RecordChange"
    app.Render
    Set probe = app.Component("trk")
    Set thumbShape = host.Shapes("rdm_wid17_trk__thumb")

    ' A visible slider on the active sheet demands the pump for its press
    ' watch.
    transcript = "dragWatchDemandsPump=" & CStr(ReDimUI.HasPendingWork)

    probe.BeginSlideDrag app, 0.3
    transcript = transcript & "|dragging=" & CStr(probe.IsSlideDragging)
    transcript = transcript & "|pressValue=" & app.State("level")
    transcript = transcript & "|thumbAccent=" & _
        CStr(thumbShape.Fill.ForeColor.RGB = app.Theme.PrimaryColor)

    probe.ContinueSlideDrag app, 0.9
    transcript = transcript & "|liveValue=" & app.State("level")
    transcript = transcript & "|noChangeDuringHold=" & CStr(gChangeCount = 0)

    probe.EndSlideDrag app
    transcript = transcript & "|releasedFiredChange=" & CStr(gChangeCount = 1)
    transcript = transcript & "|thumbWhiteAgain=" & _
        CStr(thumbShape.Fill.ForeColor.RGB = RGB(255, 255, 255))

    ' The release click Excel delivers right after is swallowed.
    ReDimUI.DispatchShape "rdm_wid17_trk"
    transcript = transcript & "|releaseClickSwallowed=" & _
        CStr(app.State("level") = 90 And gChangeCount = 1)

    ' A no-movement session fires nothing.
    Sleep 500
    probe.BeginSlideDrag app, 0.9
    probe.EndSlideDrag app
    transcript = transcript & "|noMoveNoChange=" & CStr(gChangeCount = 1)
    ReDimUI.AutoPump True
    TestSlideDrag = transcript
End Function

' Physical ground truth for the click mapping: a fixed screen pixel must
' map to absolute points that shift by exactly the scroll delta, and point
' spans must shrink with zoom. Catches both a wrong conversion factor and
' a wrong input-model choice, whatever contract the Excel build uses.
Public Function TestSlideMapping() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim probe As ReDimUI
    Dim fixedPx As Long
    Dim p1 As Double, p2 As Double
    Dim scrollDeltaPts As Double
    Dim span100 As Double, span150 As Double
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid16")
    app.SlideBar("map").AtRect 24, 24, 200, 18
    app.Render
    Set probe = app.Component("map")

    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    ActiveWindow.Zoom = 100
    fixedPx = CLng(ActiveWindow.PointsToScreenPixelsX(0)) + 500
    p1 = probe.ScreenXToSheetPoints(fixedPx)
    span100 = probe.ScreenXToSheetPoints(fixedPx + 300) - p1

    ActiveWindow.ScrollColumn = 15
    scrollDeltaPts = ActiveWindow.VisibleRange.Left
    p2 = probe.ScreenXToSheetPoints(fixedPx)
    transcript = "scrollShiftMatches=" & _
        CStr(Abs((p2 - p1) - scrollDeltaPts) < 1)

    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 150
    span150 = probe.ScreenXToSheetPoints(fixedPx + 300) - _
        probe.ScreenXToSheetPoints(fixedPx)
    transcript = transcript & "|zoomScales=" & _
        CStr(Abs(span150 - span100 / 1.5) < span100 * 0.03)
    transcript = transcript & "|span100Sane=" & _
        CStr(span100 > 100 And span100 < 400)
    ActiveWindow.Zoom = 100
    RdxStopPump
    TestSlideMapping = transcript
End Function

Public Function TestItemApi() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim picker As ReDimUI
    Dim radio As ReDimUI
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid18")
    app.SelectBox("pick").AtRect 24, 24, 130, 24
    app.SelectBox("pick").Text("pick one").Items("Alpha", "Beta", "Gamma") _
        .Value 2
    app.RadioGroup("rad").AtRect 200, 24, 130, 60
    app.RadioGroup("rad").Items("One", "Two", "Three").Value 3
    app.Render
    Set picker = app.Component("pick")
    Set radio = app.Component("rad")

    ' Insert before the selection: the selected item stays selected.
    picker.AddItem "Zeta", 1
    transcript = "countAfterInsert=" & picker.ItemCount
    transcript = transcript & "|insertedFirst=" & picker.ItemTextAt(1)
    transcript = transcript & "|selectionFollows=" & _
        CStr(picker.CurrentText = "Beta")

    ' Remove before the selection: index shifts, item keeps selection.
    picker.RemoveItem 1
    transcript = transcript & "|selectionStillBeta=" & _
        CStr(picker.CurrentText = "Beta" And picker.CurrentValue = 2)

    ' Remove the selected item by text: selection clears to placeholder.
    picker.RemoveItem "Beta"
    transcript = transcript & "|clearedToPlaceholder=" & _
        CStr(picker.CurrentValue = 0 And picker.CurrentText = "pick one")
    transcript = transcript & "|faceShowsPlaceholder=" & _
        CStr(host.Shapes("rdm_wid18_pick").TextFrame2.TextRange.Text = _
            "pick one")

    ' Replace from an array, then from a worksheet range.
    picker.ItemsFrom Array("North", "South", "East", "West")
    transcript = transcript & "|fromArray=" & picker.ItemCount
    host.Range("H1").Value = "Red"
    host.Range("H2").Value = "Green"
    host.Range("H3").Value = "Blue"
    picker.ItemsFrom host.Range("H1:H4")
    transcript = transcript & "|fromRangeSkipsBlank=" & picker.ItemCount
    transcript = transcript & "|rangeSecond=" & picker.ItemTextAt(2)

    ' Cleared list: opening shows no options.
    picker.ClearItems
    ReDimUI.DispatchShape "rdm_wid18_pick"
    transcript = transcript & "|clearedNoOptions=" & _
        CStr(Not ShapeExists(host, "rdm_wid18_pick__opt1"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid18_pick"

    ' Radio shrink: stale third-row parts are swept, selection clears.
    radio.ItemsFrom Array("Left", "Right")
    transcript = transcript & "|radioShrunk=" & radio.ItemCount
    transcript = transcript & "|radioStaleGone=" & _
        CStr(Not ShapeExists(host, "rdm_wid18_rad__c3") And _
             Not ShapeExists(host, "rdm_wid18_rad__t3"))
    transcript = transcript & "|radioSelectionCleared=" & _
        CStr(radio.CurrentValue = 0)
    ReDimUI.AutoPump True
    TestItemApi = transcript
End Function

Public Function TestComboBox() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim combo As ReDimUI
    Dim eventsWereOn As Boolean
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid20")
    app.ComboBox("color").At("E3").WritesTo("color") _
        .OnChange "TestReDimWidgets.RecordChange"
    app.ComboBox("color").Items "Red", "Green", "Gray", "Blue"
    app.Render
    Set combo = app.Component("color")

    transcript = "frameAndCaret=" & _
        CStr(ShapeExists(host, "rdm_wid20_color") And _
             ShapeExists(host, "rdm_wid20_color__caret"))

    ' Caret with empty text opens everything.
    ReDimUI.DispatchShape "rdm_wid20_color__caret"
    transcript = transcript & "|openAll=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt4"))
    transcript = transcript & "|optText=" & _
        host.Shapes("rdm_wid20_color__opt2").TextFrame2.TextRange.Text

    ' Picking writes the cell, the state, fires OnChange, and closes.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__opt2"
    transcript = transcript & "|pickedCell=" & combo.InputValue
    transcript = transcript & "|pickedState=" & app.State("color")
    transcript = transcript & "|pickChangeRan=" & gChangeCount
    transcript = transcript & "|closedAfterPick=" & _
        CStr(Not ShapeExists(host, "rdm_wid20_color__opt1"))
    transcript = transcript & "|pickedIndex=" & combo.CurrentValue

    ' Programmatic text plus caret: the list opens filtered.
    combo.InputValue = "Bl"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__caret"
    transcript = transcript & "|filteredCount=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt1") And _
             Not ShapeExists(host, "rdm_wid20_color__opt2"))
    transcript = transcript & "|filteredText=" & _
        host.Shapes("rdm_wid20_color__opt1").TextFrame2.TextRange.Text

    ' A real Enter commit with partial text auto-drops the suggestions.
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = True
    host.Range("E3").Value = "gr"
    Application.EnableEvents = eventsWereOn
    transcript = transcript & "|suggestOpened=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt2") And _
             Not ShapeExists(host, "rdm_wid20_color__opt3"))
    transcript = transcript & "|freeTextState=" & app.State("color")

    ' Exact-match commit takes the item and closes.
    Application.EnableEvents = True
    host.Range("E3").Value = "red"
    Application.EnableEvents = eventsWereOn
    transcript = transcript & "|exactClosed=" & _
        CStr(Not ShapeExists(host, "rdm_wid20_color__opt1"))
    transcript = transcript & "|exactIndex=" & combo.CurrentValue

    ' No-match commit stays closed as free text.
    Application.EnableEvents = True
    host.Range("E3").Value = "zzz"
    Application.EnableEvents = eventsWereOn
    transcript = transcript & "|noMatchClosed=" & _
        CStr(Not ShapeExists(host, "rdm_wid20_color__opt1"))
    transcript = transcript & "|noMatchState=" & app.State("color")
    ReDimUI.AutoPump True
    TestComboBox = transcript
End Function

' Cell-free fields: focus on click, characters through the key layer,
' live combo filtering per keystroke, Enter commit, Escape revert, and
' key release on blur. RdxKeyChar is the same entry OnKey drives.
Public Function TestFloatField() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid21")
    app.TextInput("name").AtRect 24, 24, 150, 20
    app.TextInput("name").WritesTo("who").OnChange _
        "TestReDimWidgets.RecordChange"
    app.ComboBox("color").AtRect 24, 60, 150, 20
    app.ComboBox("color").Items("Red", "Green", "Gray", "Blue") _
        .WritesTo "hue"
    app.Render
    Set fieldShape = host.Shapes("rdm_wid21_name")

    ' Click focuses; the face shows the insertion bar and the focus ring.
    ReDimUI.DispatchShape "rdm_wid21_name"
    transcript = "focused=" & _
        CStr(ReDimUI.FocusedComponentId = "name")
    transcript = transcript & "|caretShown=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "|")
    transcript = transcript & "|focusRing=" & _
        CStr(fieldShape.Line.ForeColor.RGB = app.Theme.PrimaryColor)
    ' The face is surface-filled, so the ink must be surface ink; the
    ' implicit primary variant would paint white-on-white and typing
    ' would be invisible even though the buffer works.
    transcript = transcript & "|inkOnSurface=" & _
        CStr(fieldShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            app.Theme.OnSurfaceColor)
    transcript = transcript & "|roundedFace=" & _
        CStr(fieldShape.AutoShapeType = msoShapeRoundedRectangle)
    transcript = transcript & "|roundedCombo=" & _
        CStr(host.Shapes("rdm_wid21_color").AutoShapeType = _
            msoShapeRoundedRectangle)

    RdxKeyChar "H"
    RdxKeyChar "i"
    transcript = transcript & "|typed=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "Hi|")
    RdxKeyChar "{BS}"
    RdxKeyChar "e"
    RdxKeyChar "y"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|committed=" & app.State("who")
    transcript = transcript & "|changeRan=" & gChangeCount
    transcript = transcript & "|blurred=" & CStr(Not ReDimUI.HasKeyboardFocus)
    transcript = transcript & "|plainText=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "Hey")

    ' Escape reverts and fires nothing.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_name"
    RdxKeyChar "x"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|reverted=" & _
        CStr(app.TextInput("name").InputValue = "Hey" And gChangeCount = 1)

    ' Float combo: focus opens, keys filter live, pick commits and blurs.
    ReDimUI.DispatchShape "rdm_wid21_color"
    transcript = transcript & "|comboOpenAll=" & _
        CStr(ShapeExists(host, "rdm_wid21_color__opt4"))
    RdxKeyChar "g"
    transcript = transcript & "|liveFiltered=" & _
        CStr(ShapeExists(host, "rdm_wid21_color__opt2") And _
             Not ShapeExists(host, "rdm_wid21_color__opt3"))
    RdxKeyChar "r"
    RdxKeyChar "e"
    transcript = transcript & "|narrowed=" & _
        CStr(ShapeExists(host, "rdm_wid21_color__opt1") And _
             Not ShapeExists(host, "rdm_wid21_color__opt2"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_color__opt1"
    transcript = transcript & "|picked=" & app.State("hue")
    transcript = transcript & "|pickBlurred=" & _
        CStr(Not ReDimUI.HasKeyboardFocus)
    transcript = transcript & "|pickClosed=" & _
        CStr(Not ShapeExists(host, "rdm_wid21_color__opt1"))

    ' Outside-press blur through the watch seam commits the field.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_name"
    RdxKeyChar "o"
    app.TextInput("name").BlurField app, True
    transcript = transcript & "|outsideCommit=" & app.State("who")
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFloatField = transcript
End Function

Public Function TestModalConfirm() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gConfirmRan = 0
    gCancelRan = 0
    Set app = ReDimUI.Mount(host, "wid6")
    app.Button("del").At("B2:C3").Text("Delete").Danger
    app.Render

    app.Confirm "Delete rows", "Remove 42 rows?", _
        "TestReDimWidgets.RecordConfirm", "TestReDimWidgets.RecordCancelChoice"
    transcript = "overlayShown=" & _
        CStr(host.Shapes("rdm_wid6_mdl_ov").Visible = msoTrue)
    transcript = transcript & "|cardText=" & _
        CStr(InStr(host.Shapes("rdm_wid6_mdl_card").TextFrame2.TextRange.Text, _
            "Remove 42 rows?") > 0)
    transcript = transcript & "|cancelHasBorder=" & _
        CStr(host.Shapes("rdm_wid6_mdl_cancel").Line.Visible = msoTrue And _
            host.Shapes("rdm_wid6_mdl_cancel").Line.ForeColor.RGB = _
            app.Theme.BorderColor)
    transcript = transcript & "|overlayCoversOrigin=" & _
        CStr(host.Shapes("rdm_wid6_mdl_ov").Left = 0 And _
            host.Shapes("rdm_wid6_mdl_ov").Top = 0)

    ReDimUI.DispatchShape "rdm_wid6_mdl_ok"
    transcript = transcript & "|confirmRan=" & gConfirmRan
    transcript = transcript & "|overlayHidden=" & _
        CStr(host.Shapes("rdm_wid6_mdl_ov").Visible = msoFalse)

    ' Reopen and cancel.
    app.Confirm "Delete rows", "Remove 42 rows?", _
        "TestReDimWidgets.RecordConfirm", "TestReDimWidgets.RecordCancelChoice"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid6_mdl_cancel"
    transcript = transcript & "|cancelRan=" & gCancelRan
    transcript = transcript & "|confirmStillOne=" & CStr(gConfirmRan = 1)
    transcript = transcript & "|overlayHiddenAgain=" & _
        CStr(host.Shapes("rdm_wid6_mdl_ov").Visible = msoFalse)
    TestModalConfirm = transcript
End Function
