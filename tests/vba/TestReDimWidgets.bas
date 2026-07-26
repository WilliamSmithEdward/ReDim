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

Public Function TestFormControls() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim checkboxShape As Shape
    Dim dropdownShape As Shape
    Dim sliderShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid3")
    app.Checkbox("chk").At("B2").Text("Enable audit").Checked(True).WritesTo "audit"
    app.Dropdown("dd").At("B4:C4").Items("North", "South", "East", "West") _
        .Value(2).WritesTo "region"
    app.Slider("sld").At("B6:E6").SliderRange(0, 200, 5).Value(80).WritesTo "volume"
    app.Render

    Set checkboxShape = host.Shapes("rdm_wid3_chk")
    Set dropdownShape = host.Shapes("rdm_wid3_dd")
    Set sliderShape = host.Shapes("rdm_wid3_sld")
    transcript = "chkChecked=" & _
        CStr(checkboxShape.ControlFormat.Value = xlOn)
    transcript = transcript & "|chkNativeCaptionEmpty=" & _
        CStr(LenB(checkboxShape.TextFrame.Characters.Text) = 0)
    transcript = transcript & "|chkCaption=" & _
        host.Shapes("rdm_wid3_chk__lbl").TextFrame2.TextRange.Text
    transcript = transcript & "|chkCaptionSize=" & _
        host.Shapes("rdm_wid3_chk__lbl").TextFrame2.TextRange.Font.Size
    transcript = transcript & "|ddItems=" & _
        dropdownShape.ControlFormat.ListCount
    transcript = transcript & "|ddIndex=" & _
        dropdownShape.ControlFormat.ListIndex
    transcript = transcript & "|sldValue=" & sliderShape.ControlFormat.Value

    ' Simulate user interaction: change control values then dispatch, the
    ' same order Excel uses when a form control fires OnAction.
    checkboxShape.ControlFormat.Value = xlOff
    ReDimUI.DispatchShape "rdm_wid3_chk"
    transcript = transcript & "|auditState=" & CStr(app.State("audit"))
    Sleep 200
    dropdownShape.ControlFormat.ListIndex = 4
    ReDimUI.DispatchShape "rdm_wid3_dd"
    transcript = transcript & "|regionState=" & app.State("region")
    Sleep 200
    sliderShape.ControlFormat.Value = 145
    ReDimUI.DispatchShape "rdm_wid3_sld"
    transcript = transcript & "|volumeState=" & app.State("volume")

    ' Clicking the themed caption part must toggle the box itself.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid3_chk__lbl"
    transcript = transcript & "|captionClickChecked=" & CStr(app.State("audit"))
    transcript = transcript & "|captionClickNative=" & _
        CStr(checkboxShape.ControlFormat.Value = xlOn)
    TestFormControls = transcript
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
