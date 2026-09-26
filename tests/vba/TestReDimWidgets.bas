Attribute VB_Name = "TestReDimWidgets"
Option Explicit

' Live scenarios for the widget set: progress, toggle, form controls, cell
' inputs, toasts, and the shapes-based modal.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function GetCaretBlinkTime Lib "user32" () As Long
Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private gChangeCount As Long
Private gConfirmRan As Long
Private gInputCount As Long
Private gLastInput As String
Private gCancelRan As Long
Private gCommandCount As Long
Private gLastCommand As String
Private gToastClicks As Long
Private gToastSender As String
Private gSenderCount As Long
Private gLastSender As String

Private Function NewCanvas() As Worksheet
    Set NewCanvas = ActiveWorkbook.Worksheets.Add
End Function

Public Sub RecordChange()
    gChangeCount = gChangeCount + 1
End Sub

Public Sub RecordConfirm()
    gConfirmRan = gConfirmRan + 1
End Sub

' Any control's click handler: counts the clicks and names the sender.
Public Sub RecordSender()
    gSenderCount = gSenderCount + 1
    gLastSender = ReDimUI.SenderId
End Sub

' A toast's click handler: the sender is the toast.
Public Sub RecordToastClick()
    gToastClicks = gToastClicks + 1
    gToastSender = ReDimUI.SenderId
End Sub

' A menu command's handler: the sender is the menu.
Public Sub RecordCommand()
    gCommandCount = gCommandCount + 1
    gLastCommand = ReDimUI.Sender.LastCommand
End Sub

' An app command's handler: the sender app names it.
Public Sub RecordAppCommand()
    gCommandCount = gCommandCount + 1
    gLastCommand = "app:" & ReDimUI.SenderApp.LastCommand
End Sub

' A table's row-open handler: the sender's value is the row.
Public Sub RecordRowOpen()
    gCommandCount = gCommandCount + 1
    gLastCommand = CStr(ReDimUI.Sender.CurrentValue)
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
    Dim frameNo As Long

    ' The knob glides over pump frames, driven here by hand.
    ReDimUI.AutoPump False
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
    For frameNo = 1 To 10
        ReDimUI.PumpOnce
    Next frameNo
    transcript = transcript & "|knobMoved=" & CStr(knob.Left > leftBefore)
    transcript = transcript & "|checkedProp=" & _
        CStr(app.Toggle("tgl").IsChecked)
    ' Clicking the knob part resolves to the same component.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid2_tgl__knob"
    transcript = transcript & "|knobClickTogglesOff=" & _
        CStr(app.State("darkMode") = False)
    ReDimUI.AutoPump True
    TestToggle = transcript
End Function

Public Function TestSelectBox() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim faceShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    ReDimUI.AutoPump False
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
        RowItem(host, "rdm_wid7_pick__opt3")

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick__opt3"
    transcript = transcript & "|pickedState=" & app.State("region")
    transcript = transcript & "|pickedFace=" & _
        faceShape.TextFrame2.TextRange.Text
    transcript = transcript & "|closedAfterPick=" & _
        CStr(Not ShapeExists(host, "rdm_wid7_pick__opt1"))
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Re-picking the current item closes the list and changes nothing.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick__opt3"
    transcript = transcript & "|repickQuiet=" & _
        CStr(gChangeCount = 1 And _
            Not ShapeExists(host, "rdm_wid7_pick__opt1"))

    ' Reopen and close by clicking the face again.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid7_pick"
    transcript = transcript & "|toggleClosed=" & _
        CStr(Not ShapeExists(host, "rdm_wid7_pick__opt1"))
    ReDimUI.AutoPump True
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

    ' The frame covers the anchor cell and takes the click; the click
    ' selects the cell so typing lands in the grid, and fires nothing.
    host.Range("A1").Select
    ReDimUI.DispatchShape "rdm_wid4_name"
    transcript = transcript & "|clickSelectsCell=" & _
        CStr(Selection.Address = "$C$3")
    transcript = transcript & "|clickNoChange=" & CStr(gChangeCount = 0)

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
    ' The slides and fades under test play only with full motion.
    ReDimUI.ReduceMotion False
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
    ' Expiry starts a fade instead of popping the shape out of existence;
    ' removal and compaction follow when the fade completes.
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|fadesOut=" & _
        CStr(ShapeExists(host, toastName) And _
            host.Shapes(toastName).Fill.Transparency > 0.05)
    Dim fadeTicks As Long
    For fadeTicks = 1 To 7
        ReDimUI.PumpOnce
    Next fadeTicks
    transcript = transcript & "|removedAfterTtl=" & _
        CStr(Not ShapeExists(host, toastName))

    ' Click dismissal: the click starts the same fade-out.
    Set toastValue = app.Toast("Click me.", 60000)
    toastName = "rdm_wid5_" & toastValue.ComponentId
    ReDimUI.DispatchShape toastName
    For fadeTicks = 1 To 8
        ReDimUI.PumpOnce
    Next fadeTicks
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
    ReDimUI.ReduceMotion
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

' True for a list pager drawn with nothing beyond it: the arrow alone.
Private Function PagerInert(ByVal host As Worksheet, ByVal shapeName As String) As Boolean
    If Not ShapeExists(host, shapeName) Then Exit Function
    PagerInert = (Len(host.Shapes(shapeName).TextFrame2.TextRange.Text) = 1)
End Function

' "1" when the probe button takes keyCode as its shortcut, "0" when
' Shortcut refuses it.
Private Function ShortcutAccepts(ByVal app As ReDimUI, ByVal keyCode As String) As String
    On Error Resume Next
    app.Button("probe").Shortcut keyCode
    ShortcutAccepts = IIf(Err.Number = 0, "1", "0")
    Err.Clear
    On Error GoTo 0
End Function

' What the workbook name that outlives a reset lists of ReDim's bound
' keys, or nothing when there is no such name.
Private Function RecordedKeys() As String
    On Error Resume Next
    RecordedKeys = ThisWorkbook.Names("rdm_bound_keys").RefersTo
    On Error GoTo 0
End Function

' The text color a drawn shape shows.
Private Function InkOf(ByVal host As Worksheet, ByVal shapeName As String) As Long
    InkOf = host.Shapes(shapeName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB
End Function

' Empties a focused float field through the keys, one Backspace per
' character, as a user would.
Private Sub BackspaceAll(ByVal fieldValue As ReDimUI)
    Dim keyNo As Long

    For keyNo = 1 To Len(fieldValue.InputValue)
        RdxKeyChar "{BS}"
    Next keyNo
End Sub

' A list row's item, after the check gutter and its tab.
Private Function RowItem(ByVal host As Worksheet, ByVal shapeName As String) As String
    Dim rowText As String

    rowText = host.Shapes(shapeName).TextFrame2.TextRange.Text
    RowItem = Mid$(rowText, InStr(1, rowText, vbTab) + 1)
End Function

' True when a list row's gutter carries the check.
Private Function RowChecked(ByVal host As Worksheet, ByVal shapeName As String) As Boolean
    RowChecked = (Left$(host.Shapes(shapeName).TextFrame2.TextRange.Text, 2) _
        = ChrW(10003) & vbTab)
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
    ' The slides and fades under test play only with full motion.
    ReDimUI.ReduceMotion False
    ReDimUI.ResetTickFaults
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
    ' The close button rides the slide and settles with the card.
    transcript = transcript & "|closeRides=" & _
        CStr(Abs(host.Shapes(firstName & "__tx").Top - _
            (host.Shapes(firstName).Top + 3)) < 0.1)

    Set secondToast = app.Toast("two", 60000)
    For ticks = 1 To 10
        ReDimUI.PumpOnce
    Next ticks
    firstSettledTop = host.Shapes(firstName).Top
    secondSettledTop = host.Shapes("rdm_wid8_" & secondToast.ComponentId).Top
    transcript = transcript & "|secondBelowFirst=" & _
        CStr(secondSettledTop > firstSettledTop)

    ' Dismissing the first toast fades it, then compacts the stack: the
    ' survivor slides up into slot one.
    ReDimUI.DispatchShape firstName
    For ticks = 1 To 18
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

    ' The rail is sticky while toasts live: scrolling between spawns moved
    ' the recomputed viewport-clamped origin, and a newcomer placed from
    ' the fresh origin could land on top of live toasts. A newcomer must
    ' join the existing column exactly one slot pitch below.
    Dim thirdShape As Shape
    Dim fourthToast As ReDimUI
    Dim fourthShape As Shape
    Dim priorScrollRow As Long
    Set thirdShape = host.Shapes("rdm_wid8_" & thirdToast.ComponentId)
    priorScrollRow = ActiveWindow.ScrollRow
    ActiveWindow.ScrollRow = 40
    Set fourthToast = app.Toast("four", 60000)
    For ticks = 1 To 10
        ReDimUI.PumpOnce
    Next ticks
    Set fourthShape = host.Shapes("rdm_wid8_" & fourthToast.ComponentId)
    transcript = transcript & "|railSticksLeft=" & _
        CStr(Abs(fourthShape.Left - thirdShape.Left) < 0.1)
    transcript = transcript & "|railSticksTop=" & _
        CStr(Abs((fourthShape.Top - thirdShape.Top) - 46) < 0.1)
    ActiveWindow.ScrollRow = priorScrollRow

    ' A newcomer spawned mid-glide must enter BELOW the column as drawn:
    ' its slot is model-correct, but a survivor easing upward may still
    ' visually occupy it. Dismiss the slot-one toast, pump exactly into
    ' the survivors' glide, spawn - the newcomer starts one full pitch
    ' under the lowest gliding survivor, never on top of it.
    Dim fifthToast As ReDimUI
    Dim fifthShape As Shape
    ReDimUI.DispatchShape "rdm_wid8_" & secondToast.ComponentId
    For ticks = 1 To 6
        ReDimUI.PumpOnce
    Next ticks
    Set fifthToast = app.Toast("five", 60000)
    Set fifthShape = host.Shapes("rdm_wid8_" & fifthToast.ComponentId)
    transcript = transcript & "|glideEntryBelow=" & _
        CStr(fifthShape.Top - fourthShape.Top >= 45.5)
    For ticks = 1 To 16
        ReDimUI.PumpOnce
    Next ticks
    transcript = transcript & "|settledPitchA=" & _
        CStr(Abs((fourthShape.Top - thirdShape.Top) - 46) < 0.1)
    transcript = transcript & "|settledPitchB=" & _
        CStr(Abs((fifthShape.Top - fourthShape.Top) - 46) < 0.1)
    ' The pump swallows faults silently by design; this proves the
    ' silence across the scenario's dozens of forced ticks was earned.
    transcript = transcript & "|tickFaults=" & ReDimUI.TickFaultCount
    ReDimUI.ReduceMotion
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
    app.Label("info").AtRect(24, 24, 60, 20).Text("Info")
    app.Render

    ' Default rail: just outside the content's right edge. The content is
    ' narrow so the rail fits even a small Excel window, which would
    ' clamp it into view.
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

    ' A stepper's press watch arms the pump at Render; the pump stays off
    ' so no timer outlives the harness call.
    ReDimUI.AutoPump False
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
    ReDimUI.AutoPump True
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
    transcript = "partsExist=" & CStr(Not trackShape Is Nothing _
        And Not fillShape Is Nothing And Not thumbShape Is Nothing)
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

    probe.EndSlideDrag
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
    probe.EndSlideDrag
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
    p1 = probe.ScreenToSheetPoints(fixedPx, False)
    span100 = probe.ScreenToSheetPoints(fixedPx + 300, False) - p1

    ActiveWindow.ScrollColumn = 15
    scrollDeltaPts = ActiveWindow.VisibleRange.Left
    p2 = probe.ScreenToSheetPoints(fixedPx, False)
    transcript = "scrollShiftMatches=" & _
        CStr(Abs((p2 - p1) - scrollDeltaPts) < 1)

    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 150
    span150 = probe.ScreenToSheetPoints(fixedPx + 300, False) - _
        probe.ScreenToSheetPoints(fixedPx, False)
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
    Dim seq As ROneCOne
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

    ' Replace from a ROneCOne sequence: a ListOf feeds a picker like any
    ' other source, and passing an object never trips the array test.
    Set seq = ROneCOne.Json.Deserialize("[""Mercury"",""Venus"",""Earth""]")
    On Error Resume Next
    Err.Clear
    picker.ItemsFrom seq
    transcript = transcript & "|seqErrClean=" & CStr(Err.Number = 0)
    On Error GoTo 0
    transcript = transcript & "|fromSequence=" & picker.ItemCount
    transcript = transcript & "|sequenceSecond=" & picker.ItemTextAt(2)

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

    ' The face covers the anchor cell: a click selects the cell for
    ' typing as well as opening the list, and a second click closes it.
    host.Range("A1").Select
    ReDimUI.DispatchShape "rdm_wid20_color"
    transcript = transcript & "|faceSelectsCell=" & _
        CStr(Selection.Address = "$E$3" And _
            ShapeExists(host, "rdm_wid20_color__opt1"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color"
    Sleep 200

    ' Caret with empty text opens everything.
    ReDimUI.DispatchShape "rdm_wid20_color__caret"
    transcript = transcript & "|openAll=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt4"))
    transcript = transcript & "|optText=" & _
        RowItem(host, "rdm_wid20_color__opt2")

    ' Picking writes the cell, the state, fires OnChange, and closes.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__opt2"
    transcript = transcript & "|pickedCell=" & combo.InputValue
    transcript = transcript & "|pickedState=" & app.State("color")
    transcript = transcript & "|pickChangeRan=" & gChangeCount
    transcript = transcript & "|closedAfterPick=" & _
        CStr(Not ShapeExists(host, "rdm_wid20_color__opt1"))
    transcript = transcript & "|pickedIndex=" & combo.CurrentValue

    ' The cell now names the pick, so reopening shows every item rather
    ' than trapping the list on the one match. Re-picking the item the
    ' cell already holds closes the list and changes nothing.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__caret"
    transcript = transcript & "|reopenAll=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt4"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__opt2"
    transcript = transcript & "|repickQuiet=" & _
        CStr(gChangeCount = 1 And _
            Not ShapeExists(host, "rdm_wid20_color__opt1"))

    ' Programmatic text plus caret: the list opens filtered.
    combo.InputValue = "Bl"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid20_color__caret"
    transcript = transcript & "|filteredCount=" & _
        CStr(ShapeExists(host, "rdm_wid20_color__opt1") And _
             Not ShapeExists(host, "rdm_wid20_color__opt2"))
    transcript = transcript & "|filteredText=" & _
        RowItem(host, "rdm_wid20_color__opt1")

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

    ' Refocusing after a pick shows the whole list, not just the pick,
    ' and the first edit filters again.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_color"
    transcript = transcript & "|reopenShowsAll=" & _
        CStr(ShapeExists(host, "rdm_wid21_color__opt4"))
    RdxKeyChar "{BS}"
    transcript = transcript & "|editFilters=" & _
        CStr(ShapeExists(host, "rdm_wid21_color__opt1") And _
             Not ShapeExists(host, "rdm_wid21_color__opt2"))
    ' Tab takes the suggestion the edit shows (Gree, then Green's n) and
    ' commits.
    RdxKeyChar "{TAB}"

    ' Esc on a combo closes an open list first, keeping focus and text;
    ' a second Esc reverts to the text focus found and leaves.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_color"
    RdxKeyChar "z"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escClosesList=" & _
        CStr(ReDimUI.HasKeyboardFocus And _
            app.ComboBox("color").InputValue = "Greenz" And _
            Not ShapeExists(host, "rdm_wid21_color__optn"))
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escEscReverts=" & _
        CStr(Not ReDimUI.HasKeyboardFocus And _
            app.ComboBox("color").InputValue = "Green")

    ' Outside-press blur through the watch seam commits the field.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_name"
    RdxKeyChar "o"
    app.TextInput("name").BlurField app, True
    transcript = transcript & "|outsideCommit=" & app.State("who")

    ' Tab commits exactly like Enter and moves focus to the next field.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_name"
    RdxKeyChar "x"
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabCommit=" & app.State("who")
    transcript = transcript & "|tabMovesOn=" & _
        CStr(ReDimUI.FocusedComponentId = "color")

    ' A cell click leaves a moved selection behind, and the next frame's
    ' selection poll commits on it - the press itself can be invisible
    ' to the pump (the grid's selection mouse loop holds WM_TIMER), and
    ' this works with application events off, exactly as the harness
    ' runs. Select stands in for the click; one forced tick is the frame.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid21_name"
    RdxKeyChar "z"
    host.Range("A40").Select
    RdxPumpOnce
    transcript = transcript & "|cellClickCommit=" & app.State("who")
    transcript = transcript & "|cellClickBlurred=" & _
        CStr(Not ReDimUI.HasKeyboardFocus)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFloatField = transcript
End Function

' Dual listbox: rows select, the four buttons move one or all items, the
' chosen list writes to state joined with a comma, headers count live,
' stale row shapes are swept as panels shrink.
Public Function TestTransferList() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ' A transfer list's press watch arms the pump at Render; the pump
    ' stays off so no timer outlives the harness call.
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid22")
    app.TransferList("teams").AtRect 24, 24, 360, 140
    app.TransferList("teams").ItemsFrom(Array("Alpha", "Bravo", "Echo")) _
        .ChosenFrom(Array("Charlie", "Delta")) _
        .Captions("Bench", "Roster") _
        .WritesTo("roster").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    transcript = "panels=" & CStr(ShapeExists(host, "rdm_wid22_teams") And _
        ShapeExists(host, "rdm_wid22_teams__rp"))
    transcript = transcript & "|buttons=" & _
        CStr(ShapeExists(host, "rdm_wid22_teams__mvr") And _
             ShapeExists(host, "rdm_wid22_teams__mvar") And _
             ShapeExists(host, "rdm_wid22_teams__mvl") And _
             ShapeExists(host, "rdm_wid22_teams__mval"))
    transcript = transcript & "|leftRows=" & _
        CStr(ShapeExists(host, "rdm_wid22_teams__al3") And _
             Not ShapeExists(host, "rdm_wid22_teams__al4"))
    transcript = transcript & "|rightRows=" & _
        CStr(ShapeExists(host, "rdm_wid22_teams__cl2") And _
             Not ShapeExists(host, "rdm_wid22_teams__cl3"))
    transcript = transcript & "|headerCounts=" & _
        CStr(host.Shapes("rdm_wid22_teams__hl").TextFrame2.TextRange.Text _
            = "Bench (3)" And _
            host.Shapes("rdm_wid22_teams__hr").TextFrame2.TextRange.Text _
            = "Roster (2)")

    ' Selecting a row shows the accent and fires nothing; clicking it
    ' again toggles it back out of the selection set.
    ReDimUI.DispatchShape "rdm_wid22_teams__al2"
    transcript = transcript & "|rowSelected=" & _
        CStr(host.Shapes("rdm_wid22_teams__al2").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor And _
            InkOf(host, "rdm_wid22_teams__al2") = app.Theme.OnPrimaryColor)
    ' Selection shows as a check too, not by color alone.
    transcript = transcript & "|selectedCheck=" & _
        CStr(RowChecked(host, "rdm_wid22_teams__al2") And _
            Not RowChecked(host, "rdm_wid22_teams__al1") And _
            RowItem(host, "rdm_wid22_teams__al2") = "Bravo")
    transcript = transcript & "|selectNoChange=" & CStr(gChangeCount = 0)
    ' Past the double-click time: a second click within it moves the row.
    Sleep GetDoubleClickTime() + 60
    ReDimUI.DispatchShape "rdm_wid22_teams__al2"
    transcript = transcript & "|toggledOff=" & _
        CStr(host.Shapes("rdm_wid22_teams__al2").Fill.ForeColor.RGB <> _
            app.Theme.PrimaryColor And _
            InkOf(host, "rdm_wid22_teams__al2") = app.Theme.OnSurfaceColor _
            And Not RowChecked(host, "rdm_wid22_teams__al2"))

    ' Multi-select: toggle two rows, both carry the accent, one move
    ' transfers both in list order.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__al1"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__al3"
    transcript = transcript & "|multiSelected=" & _
        CStr(host.Shapes("rdm_wid22_teams__al1").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor And _
            host.Shapes("rdm_wid22_teams__al3").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mvr"
    transcript = transcript & "|movedState=" & app.State("roster")
    transcript = transcript & "|movedCounts=" & _
        CStr(app.TransferList("teams").ItemCount = 1 And _
             app.TransferList("teams").ChosenCount = 4)
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Move all right, then all back; empty panels sweep their rows.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mvar"
    transcript = transcript & "|allRight=" & _
        CStr(app.TransferList("teams").ItemCount = 0 And _
             app.TransferList("teams").ChosenCount = 5 And _
             Not ShapeExists(host, "rdm_wid22_teams__al1"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mval"
    transcript = transcript & "|allLeft=" & _
        CStr(app.TransferList("teams").ItemCount = 5 And _
             app.TransferList("teams").ChosenCount = 0 And _
             LenB(CStr(app.State("roster"))) = 0)
    transcript = transcript & "|changeTotal=" & gChangeCount

    ' Move-one with nothing selected is a no-op.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mvr"
    transcript = transcript & "|noSelNoOp=" & _
        CStr(app.TransferList("teams").ChosenCount = 0 And gChangeCount = 3)

    ' Pull two back left from the right panel in one move.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mvar"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__cl1"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__cl2"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid22_teams__mvl"
    transcript = transcript & "|multiBack=" & _
        CStr(app.TransferList("teams").ItemCount = 2 And _
             app.TransferList("teams").ItemTextAt(1) = "Charlie" And _
             app.TransferList("teams").ItemTextAt(2) = "Delta" And _
             app.TransferList("teams").ChosenTextAt(1) = "Alpha")
    transcript = transcript & "|changeFinal=" & gChangeCount
    ReDimUI.AutoPump True
    TestTransferList = transcript
End Function

' Checkbox list: rows toggle from box or caption, the select-all header
' is tri-state with standard semantics, programmatic checks are silent,
' and checks follow their items through inserts and removals.
Public Function TestCheckList() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    Set app = ReDimUI.Mount(host, "wid23")
    app.CheckList("feat").AtRect 24, 24, 170, 100
    app.CheckList("feat").ItemsFrom(Array("Alpha", "Bravo", "Charlie")) _
        .CheckedFrom(Array("Bravo")) _
        .WritesTo("features").OnChange "TestReDimWidgets.RecordChange"
    app.Render

    transcript = "parts=" & CStr(ShapeExists(host, "rdm_wid23_feat") And _
        ShapeExists(host, "rdm_wid23_feat__b2") And _
        ShapeExists(host, "rdm_wid23_feat__b3") And _
        ShapeExists(host, "rdm_wid23_feat__t1") And _
        ShapeExists(host, "rdm_wid23_feat__mb") And _
        ShapeExists(host, "rdm_wid23_feat__mt"))
    transcript = transcript & "|headerText=" & _
        host.Shapes("rdm_wid23_feat__mt").TextFrame2.TextRange.Text
    transcript = transcript & "|seeded=" & _
        CStr(app.CheckList("feat").IsItemChecked(2) And _
            host.Shapes("rdm_wid23_feat__b2").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)
    transcript = transcript & "|mixedDash=" & _
        CStr(host.Shapes("rdm_wid23_feat__mb").TextFrame2.TextRange.Text _
            = "-")

    ' The caption is a full-row hit target: it toggles like the box.
    ReDimUI.DispatchShape "rdm_wid23_feat__t1"
    transcript = transcript & "|captionToggled=" & app.State("features")
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Select-all from mixed checks everything; from all, unchecks all.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid23_feat__mb"
    transcript = transcript & "|allChecked=" & _
        CStr(app.CheckList("feat").CheckedCount = 3 And _
            app.State("features") = "Alpha, Bravo, Charlie")
    transcript = transcript & "|masterCheckGlyph=" & _
        CStr(host.Shapes("rdm_wid23_feat__mb").TextFrame2.TextRange.Text _
            = ChrW(10003))
    ' A row that changes only its check rewrites only its box state; the
    ' box must still read fully checked or fully cleared.
    transcript = transcript & "|rowsChecked=" & _
        CStr(host.Shapes("rdm_wid23_feat__b3").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor And _
            host.Shapes("rdm_wid23_feat__b3").TextFrame2.TextRange.Text = _
            ChrW(10003) And _
            InkOf(host, "rdm_wid23_feat__b3") = app.Theme.OnPrimaryColor)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid23_feat__mt"
    transcript = transcript & "|noneChecked=" & _
        CStr(app.CheckList("feat").CheckedCount = 0 And _
            LenB(CStr(app.State("features"))) = 0)
    transcript = transcript & "|rowsCleared=" & _
        CStr(host.Shapes("rdm_wid23_feat__b3").Fill.ForeColor.RGB = _
            app.Theme.SurfaceColor And _
            LenB(host.Shapes("rdm_wid23_feat__b3").TextFrame2.TextRange.Text) _
            = 0 And host.Shapes("rdm_wid23_feat__b3").Line.Visible = msoTrue)
    transcript = transcript & "|changeAfterMaster=" & gChangeCount

    ' The main shape is row one's box.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid23_feat"
    transcript = transcript & "|mainShapeRow1=" & _
        CStr(app.State("features") = "Alpha" And gChangeCount = 4)

    ' Programmatic checks render but stay silent.
    app.CheckList("feat").SetItemChecked 3, True
    transcript = transcript & "|progSilent=" & _
        CStr(app.CheckList("feat").CheckedCount = 2 And _
            gChangeCount = 4 And app.State("features") = "Alpha")

    ' Checks follow their items through inserts and removals.
    app.CheckList("feat").AddItem "Zeta", 1
    transcript = transcript & "|insertShift=" & _
        CStr(app.CheckList("feat").IsItemChecked(2) And _
            app.CheckList("feat").IsItemChecked(4) And _
            Not app.CheckList("feat").IsItemChecked(1))
    app.CheckList("feat").RemoveItem 2
    transcript = transcript & "|removeShift=" & _
        CStr(app.CheckList("feat").CheckedCount = 1 And _
            app.CheckList("feat").IsItemChecked(3))

    ' New items at the same count keep the layout: only captions rewrite.
    app.CheckList("feat").ItemsFrom Array("Xray", "Yankee", "Zulu")
    app.Render
    transcript = transcript & "|captionSwap=" & _
        host.Shapes("rdm_wid23_feat__t1").TextFrame2.TextRange.Text & "," & _
        host.Shapes("rdm_wid23_feat__t3").TextFrame2.TextRange.Text

    ' Opting out of the header sweeps its parts.
    app.CheckList("feat").WithSelectAll False
    app.Render
    transcript = transcript & "|headerOff=" & _
        CStr(Not ShapeExists(host, "rdm_wid23_feat__mb"))
    TestCheckList = transcript
End Function

' Image control: picture fill from a file path, a themed placeholder for
' missing files, click dispatch, a state-bound source swap, and the
' embedded picture surviving its source file's deletion.
Public Function TestImage() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim imgA As String
    Dim imgB As String
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    imgA = Environ$("TEMP") & "\rdm_img_a.png"
    imgB = Environ$("TEMP") & "\rdm_img_b.png"
    ExportColorPng host, imgA, RGB(31, 111, 76)
    ExportColorPng host, imgB, RGB(180, 60, 40)

    Set app = ReDimUI.Mount(host, "wid24")
    app.Image("pic").AtRect(24, 24, 120, 80).Source(imgA) _
        .OnClick "TestReDimWidgets.RecordChange"
    app.Image("lost").AtRect(160, 24, 120, 60).Source "C:\nope\gone.png"
    app.Render

    transcript = "picFill=" & _
        CStr(host.Shapes("rdm_wid24_pic").Fill.Type = msoFillPicture)
    transcript = transcript & "|rounded=" & _
        CStr(host.Shapes("rdm_wid24_pic").AutoShapeType = _
            msoShapeRoundedRectangle)
    transcript = transcript & "|placeholder=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.Type <> msoFillPicture And _
            InStr(1, host.Shapes("rdm_wid24_lost").TextFrame2.TextRange.Text, _
            "not found") > 0)

    ' Clicks dispatch like any control.
    ReDimUI.DispatchShape "rdm_wid24_pic"
    transcript = transcript & "|clickRan=" & gChangeCount

    ' A state-bound source turns the placeholder into a picture.
    app.Image("lost").BindSource "logo"
    app.SetState "logo", imgB
    transcript = transcript & "|boundSwap=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.Type = msoFillPicture)

    ' The embedded picture survives its source file going away.
    Kill imgB
    app.SetState "logo", imgB
    transcript = transcript & "|embeddedKept=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.Type = msoFillPicture)

    ' A new source that is missing, or none, never shows the last one's
    ' picture, and the placeholder takes a new theme's colors.
    app.SetState "logo", "C:\nope\other.png"
    transcript = transcript & "|newMissingPlaceholder=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.Type <> msoFillPicture)
    app.SetState "logo", imgA
    app.SetState "logo", ""
    transcript = transcript & "|clearedPlaceholder=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.Type <> msoFillPicture)
    app.SetTheme ReDimUI.ThemeDark
    transcript = transcript & "|placeholderThemed=" & _
        CStr(host.Shapes("rdm_wid24_lost").Fill.ForeColor.RGB = ReDimUI.ThemeDark.MutedColor)
    app.SetTheme ReDimUI.ThemeLight
    On Error Resume Next
    Kill imgA
    On Error GoTo 0
    TestImage = transcript
End Function

' Fixes from the review of the table, sparkline, badge, stack, check
' list, and dialog chrome: a table's rows and headers come back whole
' after a shrink, a live sparkline stays under a dialog, a badge's hit
' test covers its words, a stretched member takes its own width back and
' a removed stack's members go free, a check list keeps its checks when
' its items are replaced, the command palette stays shut under a
' dialog, and "mdl_" ids belong to the dialogs.
Public Function TestReviewFixes() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim lineOrder As Long

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid84")
    app.Table("t").AtRect(24, 24, 300, 82).FontSize(9).Columns "Name", "Qty"
    app.Table("t").AddRow "Pear", 3
    app.Table("t").AddRow "Apple", 12
    app.Table("t").AddRow "Fig", 3
    app.Sparkline("sp").AtRect(360, 24, 120, 40).ValuesFrom Array(1, 3, 2)
    app.Badge("bd").AtRect(360, 90, 0, 18).Text "Overdue now"
    app.Stack("col").AtRect(24, 140, 200, 0).Gap(6).Stretch
    app.Label("wide").Sized(60, 18).Text("Stretched").InStack "col"
    app.Stack("hid").AtRect(260, 140, 160, 0).Gap 6
    app.Label("inHid").Sized(80, 18).Text("Hidden with it").InStack "hid"
    app.CheckList("ck").AtRect(24, 260, 160, 80).Items("A", "B", "C").CheckedFrom Array("B")
    app.CommandPalette ""
    app.Render

    app.Table("t").AtRect 24, 24, 300, 62
    app.Table("t").AtRect 24, 24, 300, 82
    transcript = "tableRowBack=" & CStr(ShapeExists(host, "rdm_wid84_t__tr3"))
    If ShapeExists(host, "rdm_wid84_t__tr3") Then
        transcript = transcript & "/" & CStr(host.Shapes("rdm_wid84_t__tr3").Top > _
            host.Shapes("rdm_wid84_t__tr2").Top)
    End If

    app.Confirm "Wait", "A dialog is up."
    app.Sparkline("sp").ValuesFrom Array(3, 1, 2)
    app.Sparkline("sp").ValuesFrom Array(3, 1, 2, 5)
    lineOrder = host.Shapes("rdm_wid84_sp__sl").ZOrderPosition
    transcript = transcript & "|sparkUnderDialog=" & CStr(lineOrder < _
        host.Shapes("rdm_wid84_mdl_ov").ZOrderPosition) & "/" & _
        host.Shapes("rdm_wid84_sp__sl").Nodes.Count
    app.OpenCommandPalette
    transcript = transcript & "|paletteShut=" & CStr(Not ShapeExists(host, _
        "rdm_wid84_mdl_pal_card"))
    app.CloseModal

    transcript = transcript & "|badgeHit=" & CStr(app.Badge("bd").CoversPoint(390, 99))

    app.Stack("col").Stretch False
    app.Render
    transcript = transcript & "|ownWidthBack=" & CStr(Abs(host.Shapes("rdm_wid84_wide").Width - _
        60) < 0.5)
    app.Stack("hid").Visible False
    app.Stack("hid").Remove
    transcript = transcript & "|memberFreed=" & CStr(host.Shapes("rdm_wid84_inHid").Visible = _
        msoTrue)

    app.CheckList("ck").Items "B", "C", "D"
    transcript = transcript & "|checksKept=" & CStr(app.CheckList("ck").IsItemChecked(1) And _
        Not app.CheckList("ck").IsItemChecked(2) And app.CheckList("ck").CheckedCount = 1)

    On Error Resume Next
    app.Label "mdl_mine"
    transcript = transcript & "|mdlReserved=" & Err.Description
    Err.Clear
    On Error GoTo 0
    app.Label("toast_1").AtRect(24, 360, 80, 18).Text "Mine"
    app.Toast "Hello"
    transcript = transcript & "|toastSkips=" & CStr(app.Label("toast_1").CurrentText = "Mine")

    app.Table("ex").AtRect(24, 400, 300, 82).Columns "Code", "Formula"
    app.Table("ex").AddRow "00123", "=1+1"
    app.Table("ex").ExportTo host.Range("K1")
    transcript = transcript & "|exportLiteral=" & CStr(VarType(host.Range("K2").Value) = _
        vbString And host.Range("K2").Value = "00123" And Not host.Range("L2").HasFormula _
        And host.Range("L2").Value = "=1+1")

    app.Table("dt").AtRect(24, 500, 300, 82).Columns "When"
    app.Table("dt").AddRow DateSerial(2026, 3, 5)
    app.Table("dt").FilterRows "Mar"
    transcript = transcript & "|formatFilter=" & app.Table("dt").ShownRowCount
    app.Table("dt").ColumnFormat 1, "mmm d"
    transcript = transcript & "/" & app.Table("dt").ShownRowCount
    app.Unmount True
    TestReviewFixes = transcript
End Function

Private Sub ExportColorPng( _
    ByVal host As Worksheet, _
    ByVal targetPath As String, _
    ByVal fillColor As Long _
)
    Dim chartHost As ChartObject

    On Error Resume Next
    Kill targetPath
    On Error GoTo 0
    Set chartHost = host.ChartObjects.Add(0, 0, 120, 80)
    chartHost.Chart.ChartArea.Format.Fill.ForeColor.RGB = fillColor
    chartHost.Chart.Export targetPath, "PNG"
    chartHost.Delete
End Sub

' Multi-line float input: Enter inserts a newline and keeps focus,
' Ctrl+Enter commits, the text anchors to the top, and single-line
' fields keep the Enter-commits convention.
Public Function TestMultiLineInput() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim faceRaw As String
    Dim transcript As String

    Set host = NewCanvas()
    gChangeCount = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid25")
    app.TextInput("notes").AtRect(24, 24, 200, 60).MultiLine _
        .WritesTo("noteText").OnChange "TestReDimWidgets.RecordChange"
    app.Render
    Set fieldShape = host.Shapes("rdm_wid25_notes")

    transcript = "topAnchored=" & _
        CStr(fieldShape.TextFrame2.VerticalAnchor = msoAnchorTop)

    ReDimUI.DispatchShape "rdm_wid25_notes"
    RdxKeyChar "a"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterStaysFocused=" & _
        CStr(ReDimUI.HasKeyboardFocus)
    RdxKeyChar "b"
    ' Shape text may normalize the newline character; normalize back
    ' before comparing.
    faceRaw = fieldShape.TextFrame2.TextRange.Text
    faceRaw = Replace(faceRaw, vbCrLf, vbLf)
    faceRaw = Replace(faceRaw, vbCr, vbLf)
    faceRaw = Replace(faceRaw, Chr$(11), vbLf)
    transcript = transcript & "|newlineInFace=" & _
        CStr(faceRaw = "a" & vbLf & "b|")
    RdxKeyChar "{CTRLENTER}"
    transcript = transcript & "|ctrlEnterCommits=" & _
        CStr(Not ReDimUI.HasKeyboardFocus And _
            app.State("noteText") = "a" & vbLf & "b")
    transcript = transcript & "|changeRan=" & gChangeCount

    ' Single-line fields keep the Enter-commits convention.
    app.TextInput("one").AtRect 24, 100, 150, 22
    app.TextInput("one").WritesTo "oneLine"
    app.Render
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid25_one"
    RdxKeyChar "x"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|singleLineEnterCommits=" & _
        CStr(Not ReDimUI.HasKeyboardFocus And app.State("oneLine") = "x")

    ' Overflow follows the caret: the focused view windows the tail of
    ' the buffer, with a leading ellipsis for trimmed content, while the
    ' buffer and the committed value stay complete. 200x60 at 11pt holds
    ' three visual lines; four hard lines window to the last three.
    Sleep 200
    app.TextInput("notes").InputValue = vbNullString
    ReDimUI.DispatchShape "rdm_wid25_notes"
    Dim lineNo As Long
    For lineNo = 1 To 4
        RdxKeyChar "L"
        RdxKeyChar CStr(lineNo)
        If lineNo < 4 Then RdxKeyChar "{ENTER}"
    Next lineNo
    faceRaw = fieldShape.TextFrame2.TextRange.Text
    faceRaw = Replace(faceRaw, vbCrLf, vbLf)
    faceRaw = Replace(faceRaw, vbCr, vbLf)
    faceRaw = Replace(faceRaw, Chr$(11), vbLf)
    transcript = transcript & "|tailWindow=" & _
        CStr(faceRaw = ChrW(8230) & "L2" & vbLf & "L3" & vbLf & "L4|")
    RdxKeyChar "{CTRLENTER}"
    transcript = transcript & "|fullCommit=" & _
        CStr(app.State("noteText") = _
            "L1" & vbLf & "L2" & vbLf & "L3" & vbLf & "L4")

    ' A long single line windows to its rightmost characters.
    Sleep 200
    app.TextInput("one").InputValue = vbNullString
    ReDimUI.DispatchShape "rdm_wid25_one"
    Dim keyNo As Long
    For keyNo = 1 To 30
        RdxKeyChar "x"
    Next keyNo
    faceRaw = host.Shapes("rdm_wid25_one").TextFrame2.TextRange.Text
    transcript = transcript & "|lineWindow=" & _
        CStr(faceRaw = ChrW(8230) & String$(20, "x") & "|")
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|lineFullCommit=" & _
        CStr(Len(CStr(app.State("oneLine"))) = 30)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestMultiLineInput = transcript
End Function

' Full caret editing: arrows move the caret, Home/End jump the line
' edges, Del deletes forward, characters insert mid-text, vertical
' moves clamp their column, and the viewport follows the caret upward
' as well as downward.
Public Function TestCaretEditing() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim faceRaw As String
    Dim transcript As String
    Dim keyNo As Long

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid26")
    app.TextInput("line").AtRect 24, 24, 150, 22
    app.TextInput("notes").AtRect(24, 60, 200, 60).MultiLine
    app.Render
    Set fieldShape = host.Shapes("rdm_wid26_line")

    ' Insert mid-text: abcd, two lefts, X.
    ReDimUI.DispatchShape "rdm_wid26_line"
    RdxKeyChar "a"
    RdxKeyChar "b"
    RdxKeyChar "c"
    RdxKeyChar "d"
    RdxKeyChar "{LEFT}"
    RdxKeyChar "{LEFT}"
    RdxKeyChar "X"
    transcript = "midInsert=" & _
        CStr(app.TextInput("line").InputValue = "abXcd" And _
            fieldShape.TextFrame2.TextRange.Text = "abX|cd")

    ' Backspace deletes before the caret, Del deletes at it.
    RdxKeyChar "{BS}"
    RdxKeyChar "{DEL}"
    transcript = transcript & "|bsAndDel=" & _
        CStr(app.TextInput("line").InputValue = "abd" And _
            fieldShape.TextFrame2.TextRange.Text = "ab|d")

    ' Home and End jump the line edges.
    RdxKeyChar "{HOME}"
    transcript = transcript & "|homeJump=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "|abd")
    RdxKeyChar "{END}"
    transcript = transcript & "|endJump=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "abd|")
    RdxKeyChar "{ESC}"

    ' Vertical moves work in hard lines and clamp the column.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid26_notes"
    RdxKeyChar "A"
    RdxKeyChar "B"
    RdxKeyChar "C"
    RdxKeyChar "D"
    RdxKeyChar "E"
    RdxKeyChar "{ENTER}"
    RdxKeyChar "X"
    RdxKeyChar "Y"
    RdxKeyChar "{UP}"
    Set fieldShape = host.Shapes("rdm_wid26_notes")
    faceRaw = NormalizedFace(fieldShape)
    transcript = transcript & "|upClampsColumn=" & _
        CStr(faceRaw = "AB|CDE" & vbLf & "XY")
    RdxKeyChar "{END}"
    faceRaw = NormalizedFace(fieldShape)
    transcript = transcript & "|endOfLine=" & _
        CStr(faceRaw = "ABCDE|" & vbLf & "XY")
    RdxKeyChar "{DOWN}"
    faceRaw = NormalizedFace(fieldShape)
    transcript = transcript & "|downClampsColumn=" & _
        CStr(faceRaw = "ABCDE" & vbLf & "XY|")
    RdxKeyChar "{ESC}"

    ' The viewport follows the caret upward: four lines in a three-line
    ' field, caret walked to the top line.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid26_notes"
    For keyNo = 1 To 4
        RdxKeyChar "L"
        RdxKeyChar CStr(keyNo)
        If keyNo < 4 Then RdxKeyChar "{ENTER}"
    Next keyNo
    RdxKeyChar "{UP}"
    RdxKeyChar "{UP}"
    RdxKeyChar "{UP}"
    faceRaw = NormalizedFace(fieldShape)
    transcript = transcript & "|upScrolled=" & _
        CStr(faceRaw = "L1|" & vbLf & "L2" & vbLf & "L3" & ChrW(8230))
    RdxKeyChar "{CTRLENTER}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestCaretEditing = transcript
End Function

Private Function NormalizedFace(ByVal fieldShape As Shape) As String
    Dim faceRaw As String

    faceRaw = fieldShape.TextFrame2.TextRange.Text
    faceRaw = Replace(faceRaw, vbCrLf, vbLf)
    faceRaw = Replace(faceRaw, vbCr, vbLf)
    faceRaw = Replace(faceRaw, Chr$(11), vbLf)
    NormalizedFace = faceRaw
End Function

' Long lists: combo and SelectBox drop lists window to eight rows (or
' ListRows) behind clickable pager rows, Up/Down walk a highlight that
' scrolls the combo window, Enter takes the highlighted match, clicked
' rows map through the scroll offset, and a pick reopens onto the whole
' list scrolled to it. Transfer panels page through scroll buttons with
' offset-mapped row selection.
Public Function TestLongLists() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim items(1 To 12) As Variant
    Dim position As Long
    Dim keyNo As Long
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid27")
    For position = 1 To 12
        items(position) = "Item" & Format$(position, "00")
    Next position
    app.ComboBox("pick").AtRect 24, 24, 150, 22
    app.ComboBox("pick").ItemsFrom(items).WritesTo "pickState"
    app.SelectBox("zone").AtRect 220, 24, 130, 24
    app.SelectBox("zone").ItemsFrom(items).WritesTo "zoneState"
    app.SelectBox("zone").Value 10
    app.TransferList("pool").AtRect 24, 240, 380, 128
    app.TransferList("pool").ItemsFrom(items).WritesTo "poolState"
    app.Render

    ' The open list windows to eight rows between two pager rows, the top
    ' one inert while nothing lies above.
    ReDimUI.DispatchShape "rdm_wid27_pick"
    transcript = "comboWindow=" & _
        CStr(ShapeExists(host, "rdm_wid27_pick__opt8") And _
            Not ShapeExists(host, "rdm_wid27_pick__opt9") And _
            ShapeExists(host, "rdm_wid27_pick__optd") And _
            PagerInert(host, "rdm_wid27_pick__optu"))
    transcript = transcript & "|moreText=" & _
        host.Shapes("rdm_wid27_pick__optd").TextFrame2.TextRange.Text

    ' Nine Downs walk the highlight past the window edge; the window
    ' scrolls and the first visible row becomes the second item.
    For keyNo = 1 To 9
        RdxKeyChar "{DOWN}"
    Next keyNo
    transcript = transcript & "|scrolledRow1=" & _
        RowItem(host, "rdm_wid27_pick__opt1")
    transcript = transcript & "|highlightLast=" & _
        CStr(host.Shapes("rdm_wid27_pick__opt8").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)

    ' Enter takes the highlighted match.
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterTakes=" & app.State("pickState")

    ' A clicked window row maps through the offset: reopen, walk down
    ' nine again, click the first visible row.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    BackspaceAll app.ComboBox("pick")
    For keyNo = 1 To 9
        RdxKeyChar "{DOWN}"
    Next keyNo
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick__opt1"
    transcript = transcript & "|clickMaps=" & app.State("pickState")

    ' Mouse paging: the bottom pager pages the window (clamped) and turns
    ' inert at the end, the top pager shows its count and pages back, and
    ' a Down after paging starts the highlight inside the visible window.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    BackspaceAll app.ComboBox("pick")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick__optd"
    transcript = transcript & "|comboPagedRow1=" & _
        RowItem(host, "rdm_wid27_pick__opt1")
    transcript = transcript & "|optuText=" & _
        host.Shapes("rdm_wid27_pick__optu").TextFrame2.TextRange.Text
    transcript = transcript & "|optdInertAtEnd=" & _
        CStr(PagerInert(host, "rdm_wid27_pick__optd"))
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|downStartsInWindow=" & _
        CStr(host.Shapes("rdm_wid27_pick__opt1").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor And _
            InkOf(host, "rdm_wid27_pick__opt1") = app.Theme.OnPrimaryColor)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick__optu"
    transcript = transcript & "|pagedBack=" & _
        CStr(RowItem(host, "rdm_wid27_pick__opt1") _
            = "Item01" And PagerInert(host, "rdm_wid27_pick__optu"))
    RdxKeyChar "{ENTER}"

    ' A pick reopens pristine: the whole list, scrolled so the pick is
    ' the first row, and Down walks on from the pick. The first edit
    ' filters again.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    For keyNo = 1 To 3
        RdxKeyChar "{DOWN}"
    Next keyNo
    RdxKeyChar "{ENTER}"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    transcript = transcript & "|reopenRow1=" & _
        RowItem(host, "rdm_wid27_pick__opt1")
    transcript = transcript & "|reopenWindowed=" & _
        CStr(ShapeExists(host, "rdm_wid27_pick__opt8") And _
            ShapeExists(host, "rdm_wid27_pick__optu") And _
            ShapeExists(host, "rdm_wid27_pick__optd"))
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|downFromPick=" & _
        CStr(host.Shapes("rdm_wid27_pick__opt2").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor And _
            host.Shapes("rdm_wid27_pick__opt1").Fill.ForeColor.RGB = _
            app.Theme.MutedColor And _
            InkOf(host, "rdm_wid27_pick__opt1") = app.Theme.OnSurfaceColor)
    RdxKeyChar "x"
    transcript = transcript & "|typingFilters=" & _
        CStr(Not ShapeExists(host, "rdm_wid27_pick__opt1"))
    BackspaceAll app.ComboBox("pick")
    RdxKeyChar "{ENTER}"

    ' A pick commits like the field's other commit paths: OnChange fires
    ' once when the pick changes the value, from the keyboard or the
    ' mouse, never again through the blur, and not at all for a re-pick
    ' of the value already there.
    app.ComboBox("pick").OnChange "TestReDimWidgets.RecordChange"
    gChangeCount = 0
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keyPickFires=" & gChangeCount
    transcript = transcript & "|keyPickState=" & app.State("pickState")
    ReDimUI.ForcePressEdge
    RdxPumpOnce
    transcript = transcript & "|keyPickNoRefire=" & gChangeCount
    gChangeCount = 0
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{UP}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keyRepickQuiet=" & gChangeCount
    gChangeCount = 0
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick__opt1"
    transcript = transcript & "|mouseRepickQuiet=" & gChangeCount
    gChangeCount = 0
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick__opt3"
    transcript = transcript & "|mousePickFires=" & gChangeCount
    transcript = transcript & "|mousePickState=" & app.State("pickState")
    app.ComboBox("pick").OnChange vbNullString
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    BackspaceAll app.ComboBox("pick")
    RdxKeyChar "{ENTER}"

    ' ListRows resizes the combo window.
    app.ComboBox("pick").ListRows 5
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pick"
    transcript = transcript & "|comboListRows=" & _
        CStr(ShapeExists(host, "rdm_wid27_pick__opt5") And _
            Not ShapeExists(host, "rdm_wid27_pick__opt6"))
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"

    ' SelectBox windowing: opening scrolls the selection into view (the
    ' window clamps at the end of the list), the top pager pages back,
    ' clicked rows map through the offset, and ListRows sizes the window.
    ReDimUI.DispatchShape "rdm_wid27_zone"
    transcript = transcript & "|selectWindow=" & _
        CStr(ShapeExists(host, "rdm_wid27_zone__opt8") And _
            Not ShapeExists(host, "rdm_wid27_zone__opt9") And _
            ShapeExists(host, "rdm_wid27_zone__optu") And _
            PagerInert(host, "rdm_wid27_zone__optd"))
    transcript = transcript & "|selectShowsPick=" & _
        RowItem(host, "rdm_wid27_zone__opt6")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_zone__optu"
    transcript = transcript & "|selectPagedBack=" & _
        CStr(RowItem(host, "rdm_wid27_zone__opt1") _
            = "Item01" And ShapeExists(host, "rdm_wid27_zone__optd"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_zone__opt2"
    transcript = transcript & "|selectRowMaps=" & app.State("zoneState")
    app.SelectBox("zone").ListRows 4
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_zone"
    transcript = transcript & "|selectListRows=" & _
        CStr(ShapeExists(host, "rdm_wid27_zone__opt4") And _
            Not ShapeExists(host, "rdm_wid27_zone__opt5") And _
            RowItem(host, "rdm_wid27_zone__opt1") _
                = "Item02")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_zone"
    On Error Resume Next
    Err.Clear
    app.TransferList("pool").ListRows 3
    transcript = transcript & "|listRowsKindGuard=" & CStr(Err.Number <> 0)
    Err.Clear
    app.SelectBox("zone").ListRows 0
    transcript = transcript & "|listRowsMinimum=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    ' Transfer panel: five rows fit, arrows page, selection maps
    ' through the offset.
    transcript = transcript & "|poolRows=" & _
        CStr(ShapeExists(host, "rdm_wid27_pool__al5") And _
            Not ShapeExists(host, "rdm_wid27_pool__al6") And _
            ShapeExists(host, "rdm_wid27_pool__ald"))
    transcript = transcript & "|pagerTarget=" & _
        CStr(host.Shapes("rdm_wid27_pool__ald").Width >= 18 And _
            host.Shapes("rdm_wid27_pool__ald").Height >= 18 And _
            host.Shapes("rdm_wid27_pool__al1").Left + _
            host.Shapes("rdm_wid27_pool__al1").Width < _
            host.Shapes("rdm_wid27_pool__ald").Left)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pool__ald"
    transcript = transcript & "|pagedRow1=" & _
        RowItem(host, "rdm_wid27_pool__al1")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pool__al2"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid27_pool__mvr"
    transcript = transcript & "|offsetMovedState=" & app.State("poolState")
    transcript = transcript & "|scrollerGoneWhenFits=" & _
        CStr(Not ShapeExists(host, "rdm_wid27_pool__cld"))
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestLongLists = transcript
End Function

' Floating chrome claims the points it covers, so pump watches cannot
' hit-test through an open drop list or a modal overlay into controls
' painted underneath.
Public Function TestChromeClaim() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim items(1 To 12) As Variant
    Dim position As Long
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "wid28")
    For position = 1 To 12
        items(position) = "Item" & Format$(position, "00")
    Next position
    app.ComboBox("pick").AtRect 24, 24, 150, 22
    app.ComboBox("pick").ItemsFrom items
    app.SelectBox("zone").AtRect 300, 24, 130, 24
    app.SelectBox("zone").ItemsFrom items
    app.SlideBar("vol").AtRect 24, 200, 200, 18
    app.SlideBar("vol").SliderRange 0, 100, 5
    app.Render

    ' A closed list claims nothing.
    transcript = "closedClaimsNothing=" & _
        CStr(Not app.PointClaimedByChrome(30, 150, "vol"))

    ' The open list claims its dropped zone, but not points beside it,
    ' and the asking component itself is excluded from the scan.
    ReDimUI.DispatchShape "rdm_wid28_pick"
    transcript = transcript & "|openClaims=" & _
        CStr(app.PointClaimedByChrome(30, 150, "vol"))
    transcript = transcript & "|besideNotClaimed=" & _
        CStr(Not app.PointClaimedByChrome(400, 150, "vol"))
    transcript = transcript & "|exceptSelf=" & _
        CStr(Not app.PointClaimedByChrome(30, 150, "pick"))
    RdxKeyChar "{ESC}"
    transcript = transcript & "|closedAgain=" & _
        CStr(Not app.PointClaimedByChrome(30, 150, "vol"))

    ' An open SelectBox claims its windowed depth - eight rows between
    ' two pagers, ending near y 290 - not all twelve items.
    ReDimUI.DispatchShape "rdm_wid28_zone"
    transcript = transcript & "|selectClaimsWindow=" & _
        CStr(app.PointClaimedByChrome(310, 250, "vol"))
    transcript = transcript & "|selectBelowWindowFree=" & _
        CStr(Not app.PointClaimedByChrome(310, 330, "vol"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid28_zone"

    ' A modal overlay claims everything it covers until it closes.
    app.Confirm "Sure?", "Chrome claim check.", vbNullString
    transcript = transcript & "|overlayClaims=" & _
        CStr(app.PointClaimedByChrome(120, 210, "vol"))
    ReDimUI.DispatchShape "rdm_wid28_mdl_cancel"
    transcript = transcript & "|overlayReleased=" & _
        CStr(Not app.PointClaimedByChrome(120, 210, "vol"))
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestChromeClaim = transcript
End Function

Public Function TestModalConfirm() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    ReDimUI.AutoPump False
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
    ReDimUI.AutoPump True
    TestModalConfirm = transcript
End Function

' Drop lists dismiss the native way: another list opening, a click on
' another control, a press off the face and rows, or a moved grid
' selection closes an open list; a press on its own rows keeps it open.
' A date picker's calendar, opened by a click and so without keyboard
' focus, closes the same ways.
Public Function TestListDismiss() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ' The forced press edge goes to the first app that ticks.
    ReDimUI.Shutdown
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid29")
    app.SelectBox("one").AtRect(24, 24, 140, 22).Items "Red", "Green", "Blue"
    app.SelectBox("two").AtRect(200, 24, 140, 22).Items "North", "South"
    app.Button("go").AtRect(24, 200, 80, 26).Text "Go"
    app.DatePicker("when").AtRect 380, 24, 150, 24
    app.Render

    ReDimUI.DispatchShape "rdm_wid29_one"
    ReDimUI.DispatchShape "rdm_wid29_two"
    transcript = "oneListPerApp=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_one__opt1") And _
            ShapeExists(host, "rdm_wid29_two__opt1"))
    ReDimUI.DispatchShape "rdm_wid29_go"
    transcript = transcript & "|clickElsewhereCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_two__opt1"))

    ' Row one of the first list spans 46 to 68 points down the sheet.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid29_one"
    ReDimUI.OverridePointer 60, 57
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|pressOnRowsKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid29_one__opt1"))
    ReDimUI.OverridePointer 420, 300
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|pressOffCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_one__opt1"))
    ReDimUI.ClearPointerOverride

    Sleep 200
    host.Range("A1").Select
    ReDimUI.DispatchShape "rdm_wid29_one"
    ReDimUI.PumpOnce
    transcript = transcript & "|selectionHeldKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid29_one__opt1"))
    host.Range("D5").Select
    ReDimUI.PumpOnce
    transcript = transcript & "|selectionMoveCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_one__opt1"))

    ' The calendar spans 380 to 574 points across and 50 to 234 down.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid29_when"
    ReDimUI.OverridePointer 450, 150
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|datePressOnKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid29_when__cb"))
    ReDimUI.OverridePointer 420, 300
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|datePressOffCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_when__cb"))
    ReDimUI.ClearPointerOverride
    Sleep 200
    host.Range("A1").Select
    ReDimUI.DispatchShape "rdm_wid29_when"
    ReDimUI.PumpOnce
    transcript = transcript & "|dateSelectionHeldKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid29_when__cb"))
    host.Range("D5").Select
    ReDimUI.PumpOnce
    transcript = transcript & "|dateSelectionMoveCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_when__cb"))

    ' A press and its release come apart, as a real click's do. A focused
    ' picker's calendar closes and its focus ends on one press off it.
    Sleep 200
    app.Component("when").Focus
    RdxKeyChar "{DOWN}"
    ReDimUI.OverridePointer 420, 300, True
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 420, 300
    ReDimUI.PumpOnce
    Sleep 200
    ReDimUI.PumpOnce
    transcript = transcript & "|focusedPressOffEnds=" & _
        CStr(Not ShapeExists(host, "rdm_wid29_when__cb") And _
            LenB(ReDimUI.FocusedComponentId) = 0)
    ' A press on another control that closed the calendar leaves nothing
    ' behind: the picker focused afterwards keeps its focus.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid29_when"
    ReDimUI.OverridePointer 60, 210, True
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    ReDimUI.DispatchShape "rdm_wid29_go"
    ReDimUI.OverridePointer 60, 210
    ReDimUI.PumpOnce
    app.Component("when").Focus
    ReDimUI.PumpOnce
    Sleep 200
    ReDimUI.PumpOnce
    transcript = transcript & "|laterFocusHolds=" & _
        CStr(ReDimUI.FocusedComponentId = "when")
    ReDimUI.ClearPointerOverride
    ReDimUI.ClearKeyboardFocus
    ReDimUI.AutoPump True
    TestListDismiss = transcript
End Function

' List conventions: the current item's row carries a check in its gutter,
' an open list with nothing to show says so in an inert row, and a list
' with no room below the visible window opens upward.
Public Function TestListConventions() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim viewBottom As Double
    Dim faceTop As Double

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid30")
    app.SelectBox("pick").AtRect(24, 24, 140, 22) _
        .Items("Red", "Green", "Blue").Value 2
    app.SelectBox("none").AtRect 200, 24, 140, 22
    app.ComboBox("find").AtRect(24, 120, 160, 22).Items "Alpha", "Beta", "Gamma"
    app.Render

    ReDimUI.DispatchShape "rdm_wid30_pick"
    transcript = "currentChecked=" & _
        CStr(RowChecked(host, "rdm_wid30_pick__opt2") And _
            Not RowChecked(host, "rdm_wid30_pick__opt1") And _
            Not RowChecked(host, "rdm_wid30_pick__opt3"))
    transcript = transcript & "|rowItem=" & RowItem(host, "rdm_wid30_pick__opt2")
    ReDimUI.DispatchShape "rdm_wid30_none"
    transcript = transcript & "|noItemsRow=" & RowItem(host, "rdm_wid30_none__optn")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid30_none__optn"
    transcript = transcript & "|emptyRowInert=" & _
        CStr(ShapeExists(host, "rdm_wid30_none__optn"))

    ' A filter that matches nothing shows the row; a match replaces it.
    ReDimUI.DispatchShape "rdm_wid30_find"
    RdxKeyChar "z"
    transcript = transcript & "|noMatchesRow=" & RowItem(host, "rdm_wid30_find__optn")
    RdxKeyChar "{BS}"
    RdxKeyChar "a"
    transcript = transcript & "|emptyRowGone=" & _
        CStr(Not ShapeExists(host, "rdm_wid30_find__optn") And _
            ShapeExists(host, "rdm_wid30_find__opt1"))
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"

    ' No room below the visible window: the list opens upward, its last
    ' row against the face and its first row on top.
    viewBottom = ActiveWindow.VisibleRange.Top + ActiveWindow.VisibleRange.Height
    faceTop = viewBottom - 30
    If faceTop < 150 Then faceTop = 150
    app.SelectBox("low").AtRect(24, faceTop, 140, 22) _
        .Items "One", "Two", "Three", "Four"
    app.Render
    ReDimUI.DispatchShape "rdm_wid30_low"
    transcript = transcript & "|opensUpward=" & _
        CStr(Abs(host.Shapes("rdm_wid30_low__opt4").Top + _
            host.Shapes("rdm_wid30_low__opt4").Height - faceTop) < 0.5)
    transcript = transcript & "|firstRowOnTop=" & _
        CStr(host.Shapes("rdm_wid30_low__opt1").Top < _
            host.Shapes("rdm_wid30_low__opt4").Top)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid30_low__opt1"
    transcript = transcript & "|upwardPick=" & CStr(app.SelectBox("low").CurrentValue = 1)
    ReDimUI.AutoPump True
    TestListConventions = transcript
End Function

' Reduced motion: toasts appear in their slot, move up at once when an
' earlier toast leaves, and leave without a fade. The caret blinks at
' the Windows rate, or not at all when Windows says not to.
Public Function TestMotionAndBlink() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim firstToast As ReDimUI
    Dim secondToast As ReDimUI
    Dim firstName As String
    Dim secondName As String
    Dim slotTop As Double
    Dim fieldShape As Shape
    Dim systemBlink As Long
    Dim tickNo As Long
    Dim transcript As String

    ReDimUI.AutoPump False
    ReDimUI.ReduceMotion True
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid31")
    app.Label("anchorlbl").At("B2").Text "x"
    app.TextInput("name").AtRect 24, 200, 160, 22
    app.Render
    transcript = "motionReduced=" & CStr(ReDimUI.MotionReduced)

    Set firstToast = app.Toast("First.", 60000)
    Set secondToast = app.Toast("Second.", 60000)
    firstName = "rdm_wid31_" & firstToast.ComponentId
    secondName = "rdm_wid31_" & secondToast.ComponentId
    slotTop = host.Shapes(firstName).Top
    transcript = transcript & "|noEntranceSlide=" & _
        CStr(Abs(host.Shapes(secondName).Top - slotTop - _
            host.Shapes(firstName).Height - 6) < 0.5)
    ReDimUI.DispatchShape firstName
    ReDimUI.PumpOnce
    transcript = transcript & "|leavesAtOnce=" & _
        CStr(Not ShapeExists(host, firstName))
    transcript = transcript & "|survivorMovesAtOnce=" & _
        CStr(Abs(host.Shapes(secondName).Top - slotTop) < 0.5)
    ReDimUI.ReduceMotion False
    transcript = transcript & "|overrideFull=" & CStr(Not ReDimUI.MotionReduced)
    ReDimUI.ReduceMotion

    ' Each forced tick is 50 ms, so the caret turns off on the first
    ' tick that reaches the Windows interval.
    systemBlink = GetCaretBlinkTime()
    ReDimUI.DispatchShape "rdm_wid31_name"
    Set fieldShape = host.Shapes("rdm_wid31_name")
    For tickNo = 1 To 60
        ReDimUI.PumpOnce
        If NormalizedFace(fieldShape) = " " Then Exit For
    Next tickNo
    If systemBlink > 0 Then
        transcript = transcript & "|blinkAtSystemRate=" & _
            CStr(tickNo = -Int(-systemBlink / 50))
    Else
        transcript = transcript & "|blinkAtSystemRate=" & CStr(tickNo > 60)
    End If
    RdxKeyChar "{ESC}"
    ReDimUI.AutoPump True
    TestMotionAndBlink = transcript
End Function

' Accessibility: each control's shape carries alternative text from its
' kind, text, and state, AltText overrides it, the contrast math matches
' WCAG, and the high-contrast theme passes every pairing.
Public Function TestAccessibility() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid32")
    app.Button("save").AtRect(24, 24, 90, 28).Text "Save"
    app.TickBox("agree").AtRect(24, 70, 160, 20).Text "Agree"
    app.Toggle("dark").AtRect 24, 100, 44, 22
    app.SelectBox("zone").AtRect(24, 140, 140, 22) _
        .Items("North", "South").Value 2
    app.ProgressBar("load").AtRect(24, 180, 160, 10).Value 40
    app.TextInput("email").AtRect(24, 230, 160, 22).Caption "Email"
    app.SelectBox("size").AtRect(220, 24, 140, 22).Items("S", "M").Placeholder "Choose a size"
    app.Render

    transcript = "buttonAlt=" & host.Shapes("rdm_wid32_save").AlternativeText
    transcript = transcript & "|captionAlt=" & host.Shapes("rdm_wid32_email").AlternativeText
    transcript = transcript & "|selectPlaceholder=" & _
        host.Shapes("rdm_wid32_size").TextFrame2.TextRange.Text
    app.Button("save").Enabled False
    transcript = transcript & "|disabledAlt=" & _
        host.Shapes("rdm_wid32_save").AlternativeText
    ReDimUI.DispatchShape "rdm_wid32_agree"
    transcript = transcript & "|tickAlt=" & _
        host.Shapes("rdm_wid32_agree").AlternativeText
    transcript = transcript & "|toggleAlt=" & _
        host.Shapes("rdm_wid32_dark").AlternativeText
    transcript = transcript & "|selectAlt=" & _
        host.Shapes("rdm_wid32_zone").AlternativeText
    transcript = transcript & "|progressAlt=" & _
        host.Shapes("rdm_wid32_load").AlternativeText
    app.Toggle("dark").AltText "Dark mode"
    transcript = transcript & "|overrideAlt=" & _
        host.Shapes("rdm_wid32_dark").AlternativeText

    transcript = transcript & "|blackOnWhite=" & _
        Format$(ReDimUI.ContrastRatio(RGB(0, 0, 0), RGB(255, 255, 255)), "0.00")
    transcript = transcript & "|sameColor=" & _
        Format$(ReDimUI.ContrastRatio(RGB(90, 90, 90), RGB(90, 90, 90)), "0.00")
    transcript = transcript & "|highContrastFails=" & _
        CStr(UBound(Split(ReDimUI.ThemeHighContrast.ContrastReport, "fail")))
    transcript = transcript & "|reportLines=" & _
        CStr(UBound(Split(ReDimUI.ThemeLight.ContrastReport, vbLf)) + 1)

    ' On high contrast's yellow track the knob turns black.
    app.SetTheme ReDimUI.ThemeHighContrast
    ReDimUI.DispatchShape "rdm_wid32_dark"
    transcript = transcript & "|knobOnAccent=" & _
        CStr(host.Shapes("rdm_wid32_dark__knob").Fill.ForeColor.RGB = _
            app.Theme.OnPrimaryColor)
    ReDimUI.AutoPump True
    TestAccessibility = transcript
End Function

' Toast conventions: a close button on every toast, a reading-time TTL
' when none is given, a tone's icon and edge, and an action button that
' dismisses the toast and runs its handler.
Public Function TestToastConventions() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim toastValue As ReDimUI
    Dim toastName As String
    Dim longText As String
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    ' Reduced motion makes each dismissal finish on one tick.
    ReDimUI.ReduceMotion True
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid33")
    app.Label("anchorlbl").At("B2").Text "x"
    app.Render

    Set toastValue = app.Toast("Saved.")
    toastName = "rdm_wid33_" & toastValue.ComponentId
    transcript = "closeButton=" & CStr(ShapeExists(host, toastName & "__tx"))
    transcript = transcript & "|plainText=" & _
        host.Shapes(toastName).TextFrame2.TextRange.Text
    transcript = transcript & "|shortTtl=" & _
        CStr(toastValue.ToastRemainingMs > 3800 And toastValue.ToastRemainingMs <= 4000)
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    transcript = transcript & "|closeDismisses=" & _
        CStr(Not ShapeExists(host, toastName) And _
            Not ShapeExists(host, toastName & "__tx"))

    longText = String$(100, "a")
    Set toastValue = app.Toast(longText)
    transcript = transcript & "|longTtl=" & _
        CStr(toastValue.ToastRemainingMs > 8800 And toastValue.ToastRemainingMs <= 9000)
    ReDimUI.DispatchShape "rdm_wid33_" & toastValue.ComponentId
    ReDimUI.PumpOnce

    Set toastValue = app.Toast("Low disk space.").Warning
    toastName = "rdm_wid33_" & toastValue.ComponentId
    transcript = transcript & "|warningIcon=" & _
        CStr(Left$(host.Shapes(toastName).TextFrame2.TextRange.Text, 2) = _
            ChrW(&H26A0) & vbTab)
    transcript = transcript & "|toneColors=" & _
        CStr(host.Shapes(toastName).Line.ForeColor.RGB = app.Theme.WarningColor _
            And host.Shapes(toastName).TextFrame2.TextRange.Characters(1, 1) _
                .Font.Fill.ForeColor.RGB = app.Theme.WarningColor _
            And InkOf(host, toastName & "__tx") = app.Theme.OnMutedColor)
    ReDimUI.DispatchShape toastName
    ReDimUI.PumpOnce

    Set toastValue = app.Toast("Row deleted.", 5000)
    toastValue.Action "Undo", "TestReDimWidgets.RecordChange"
    toastName = "rdm_wid33_" & toastValue.ComponentId
    transcript = transcript & "|actionText=" & _
        host.Shapes(toastName & "__ta").TextFrame2.TextRange.Text
    transcript = transcript & "|actionAccent=" & _
        CStr(host.Shapes(toastName & "__ta").Line.ForeColor.RGB = app.Theme.PrimaryColor)
    transcript = transcript & "|actionLongerTtl=" & _
        CStr(toastValue.ToastRemainingMs > 8800)
    ReDimUI.DispatchShape toastName & "__ta"
    ReDimUI.PumpOnce
    transcript = transcript & "|actionRan=" & CStr(gChangeCount = 1)
    transcript = transcript & "|actionDismisses=" & _
        CStr(Not ShapeExists(host, toastName & "__ta"))

    ' The close button puts a toast away at once, even with the pointer
    ' resting on it and holding its countdown, and runs no handler.
    Set toastValue = app.Toast("Row deleted.", 5000)
    toastValue.Action "Undo", "TestReDimWidgets.RecordChange"
    toastName = "rdm_wid33_" & toastValue.ComponentId
    With host.Shapes(toastName)
        ReDimUI.OverridePointer .Left + 20, .Top + .Height / 2
    End With
    ReDimUI.PumpOnce
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|closeUnderPointer=" & _
        CStr(Not ShapeExists(host, toastName) And gChangeCount = 1)

    ' A click on the card runs the toast's OnClick, with the toast as the
    ' sender, and puts it away; the close button runs nothing.
    gToastClicks = 0
    Set toastValue = app.Toast("Report ready.").OnClick("TestReDimWidgets.RecordToastClick")
    toastName = "rdm_wid33_" & toastValue.ComponentId
    With host.Shapes(toastName)
        ReDimUI.OverridePointer .Left + 20, .Top + .Height / 2
    End With
    ReDimUI.PumpOnce
    ReDimUI.DispatchShape toastName
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|cardRunsOnClick=" & CStr(gToastClicks = 1 _
        And gToastSender = toastValue.ComponentId And Not ShapeExists(host, toastName))
    ReDimUI.ClearPointerOverride
    Set toastValue = app.Toast("Report ready.").OnClick("TestReDimWidgets.RecordToastClick")
    toastName = "rdm_wid33_" & toastValue.ComponentId
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    transcript = transcript & "|closeRunsNothing=" & _
        CStr(gToastClicks = 1 And Not ShapeExists(host, toastName))
    ' A second click, landing while the toast is already on its way out,
    ' runs nothing: the first click of a double-click ran the handler.
    Set toastValue = app.Toast("Report ready.").OnClick("TestReDimWidgets.RecordToastClick")
    toastName = "rdm_wid33_" & toastValue.ComponentId
    ReDimUI.DispatchShape toastName
    Sleep 200
    ReDimUI.DispatchShape toastName
    transcript = transcript & "|secondClickRunsNothing=" & CStr(gToastClicks = 2)
    ReDimUI.PumpOnce

    ' ActionBorder colors the action button's border.
    Set toastValue = app.Toast("Moved to the archive.")
    toastValue.Action("Undo", "TestReDimWidgets.RecordChange").ActionBorder RGB(200, 30, 30)
    toastName = "rdm_wid33_" & toastValue.ComponentId
    transcript = transcript & "|actionBorder=" & _
        CStr(host.Shapes(toastName & "__ta").Line.ForeColor.RGB = RGB(200, 30, 30))
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestToastConventions = transcript
End Function

' Shortcuts: a control's Shortcut binds while its sheet is in front and
' clicks the control as a mouse click does, sender and all; a disabled
' control passes the key to the next that declares it; the tooltip and
' the alternative text name the shortcut; a key a field types is
' refused; and leaving the sheet releases the binding.
Public Function TestShortcuts() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim refused As Boolean
    Dim senderBefore As Long
    Dim tempBound As Boolean

    gSenderCount = 0
    gLastSender = vbNullString
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid73")
    app.Button("save").AtRect(24, 24, 100, 30).Text("Save").Tooltip("Saves the form") _
        .Shortcut("^s").OnClick "TestReDimWidgets.RecordSender"
    app.Button("old").AtRect(24, 70, 100, 30).Text("Old").Enabled(False).Shortcut("{F7}") _
        .OnClick "TestReDimWidgets.RecordSender"
    app.Button("new").AtRect(24, 110, 100, 30).Text("New").Shortcut("{F7}") _
        .OnClick "TestReDimWidgets.RecordSender"
    app.Render
    host.Activate
    ReDimUI.RefreshAccessKeys
    transcript = "bound=" & CStr(InStr(ReDimUI.BoundShortcuts, Chr$(1) & "^s" & Chr$(1)) > 0 _
        And InStr(ReDimUI.BoundShortcuts, Chr$(1) & "{F7}" & Chr$(1)) > 0)
    RdxShortcut "^s"
    transcript = transcript & "|clicks=" & gSenderCount & "/" & gLastSender
    RdxShortcut "{F7}"
    transcript = transcript & "|skipsDisabled=" & gSenderCount & "/" & gLastSender
    transcript = transcript & "|alt=" & _
        CStr(InStr(host.Shapes("rdm_wid73_save").AlternativeText, "Ctrl+S") > 0)

    ReDimUI.OverridePointer 74, 39
    ReDimUI.PumpOnce
    Sleep GetDoubleClickTime() + 60
    ReDimUI.PumpOnce
    transcript = transcript & "|tip=" & _
        host.Shapes("rdm_wid73_save__tt").TextFrame2.TextRange.Text
    ReDimUI.OverridePointer 600, 400
    ReDimUI.PumpOnce
    ReDimUI.ClearPointerOverride

    On Error Resume Next
    app.Button("save").Shortcut "s"
    refused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    transcript = transcript & "|refusesTypingKey=" & CStr(refused _
        And InStr(ReDimUI.BoundShortcuts, Chr$(1) & "^s" & Chr$(1)) > 0)

    ' A key that focus captured still reaches a shortcut: with a button
    ' focused, Ctrl+Z arrives as {UNDO} and clicks the control declaring it.
    app.Button("undo").AtRect(24, 150, 100, 30).Text("Undo").Shortcut("^z") _
        .OnClick "TestReDimWidgets.RecordSender"
    app.Component("new").Focus
    RdxKeyChar "{UNDO}"
    transcript = transcript & "|capturedChordClicks=" & gLastSender
    ReDimUI.ClearKeyboardFocus

    ' A shortcut goes on a control a click acts on: a slider has no one
    ' click to give, and a label with OnClick takes one.
    On Error Resume Next
    app.SlideBar("vol").AtRect(300, 24, 120, 18).Shortcut "^m"
    refused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    transcript = transcript & "|refusesSlider=" & CStr(refused)
    app.Label("help").AtRect(24, 190, 100, 20).Text("Help").Shortcut("{F1}") _
        .OnClick "TestReDimWidgets.RecordSender"
    RdxShortcut "{F1}"
    transcript = transcript & "|labelClicks=" & gLastSender

    ' Codes OnKey refuses, and Ctrl+Alt with a character, are refused.
    app.Button("probe").AtRect(300, 150, 60, 24).Text "Probe"
    transcript = transcript & "|codeRules=" & ShortcutAccepts(app, "^{") & _
        ShortcutAccepts(app, "^{Save}") & ShortcutAccepts(app, "{F16}") & _
        ShortcutAccepts(app, "^%e") & ShortcutAccepts(app, "^ ") & _
        ShortcutAccepts(app, "^%{DEL}") & ShortcutAccepts(app, "{F15}") & _
        ShortcutAccepts(app, "^{(}")
    app.Button("probe").Shortcut ""

    ' Removing a control gives its key back, and the bound keys are kept
    ' in a workbook name as well.
    app.Button("temp").AtRect(300, 190, 60, 24).Text("Temp").Shortcut "^k"
    tempBound = (InStr(ReDimUI.BoundShortcuts, Chr$(1) & "^k" & Chr$(1)) > 0)
    app.Component("temp").Remove
    transcript = transcript & "|removeReleases=" & CStr(tempBound _
        And InStr(ReDimUI.BoundShortcuts, Chr$(1) & "^k" & Chr$(1)) = 0)
    transcript = transcript & "|recorded=" & CStr(InStr(RecordedKeys(), "^s") > 0)

    ' A key pressed on a sheet no app declares it on, which a sheet change
    ' Excel did not report leaves bound, clicks nothing and goes back.
    senderBefore = gSenderCount
    NewCanvas
    RdxShortcut "^s"
    transcript = transcript & "|strayReleases=" & CStr(gSenderCount = senderBefore _
        And LenB(ReDimUI.BoundShortcuts) = 0)
    ReDimUI.RefreshAccessKeys
    transcript = transcript & "|leaveReleases=" & CStr(LenB(ReDimUI.BoundShortcuts) = 0)
    host.Activate
    ReDimUI.RefreshAccessKeys
    transcript = transcript & "|returnBinds=" & _
        CStr(InStr(ReDimUI.BoundShortcuts, Chr$(1) & "^s" & Chr$(1)) > 0)
    ReDimUI.Shutdown
    transcript = transcript & "|shutdownClearsRecord=" & CStr(LenB(RecordedKeys()) = 0)
    ReDimUI.AutoPump True
    TestShortcuts = transcript
End Function

' Toast size: a toast keeps 240 points across unless MinWidth and MaxWidth
' let it fit its words, grows as tall as its wrapped message up to
' MaxHeight, where the message ends in an ellipsis, and stacks under the
' toasts above it by their heights; when a tall toast leaves, the one
' under it moves up to its place.
Public Function TestToastSize() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim toastValue As ReDimUI
    Dim tallName As String
    Dim nextName As String
    Dim toastName As String
    Dim tallTop As Double
    Dim longWords As String
    Dim needTall As Double
    Dim edgeHost As Worksheet
    Dim edgeApp As ReDimUI
    Dim viewBottom As Double
    Dim viewRight As Double

    ReDimUI.Shutdown
    ReDimUI.AutoPump False
    ReDimUI.ReduceMotion True
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid72")
    app.Label("anchorlbl").At("B2").Text "x"
    app.Render
    longWords = "The export finished and the workbook was saved to the shared folder, " & _
        "where everyone on the team can open it and check the totals before Friday."

    Set toastValue = app.Toast("Saved.")
    toastName = "rdm_wid72_" & toastValue.ComponentId
    transcript = "shortSize=" & Format$(host.Shapes(toastName).Width, "0") & "x" & _
        Format$(host.Shapes(toastName).Height, "0")
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce

    Set toastValue = app.Toast(longWords)
    tallName = "rdm_wid72_" & toastValue.ComponentId
    With host.Shapes(tallName)
        transcript = transcript & "|growsTall=" & CStr(.Width = 240 And .Height > 40 _
            And .TextFrame2.TextRange.BoundHeight <= .Height)
        tallTop = .Top
    End With
    Set toastValue = app.Toast("Next.")
    nextName = "rdm_wid72_" & toastValue.ComponentId
    transcript = transcript & "|stacksBelow=" & CStr(host.Shapes(nextName).Top >= _
        tallTop + host.Shapes(tallName).Height + 5.5)
    ReDimUI.DispatchShape tallName & "__tx"
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|movesUp=" & _
        CStr(Abs(host.Shapes(nextName).Top - tallTop) < 0.5)
    ReDimUI.DispatchShape nextName & "__tx"
    ReDimUI.PumpOnce

    Set toastValue = app.Toast(longWords & " " & longWords).MaxHeight(70)
    toastName = "rdm_wid72_" & toastValue.ComponentId
    With host.Shapes(toastName)
        transcript = transcript & "|capsHeight=" & CStr(.Height <= 70.01 _
            And Right$(.TextFrame2.TextRange.Text, 1) = ChrW(8230))
    End With
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce

    Set toastValue = app.Toast("Saved.").MinWidth(100).MaxWidth(420)
    toastName = "rdm_wid72_" & toastValue.ComponentId
    transcript = transcript & "|fitsNarrow=" & CStr(host.Shapes(toastName).Width < 240 _
        And host.Shapes(toastName).Width >= 100)
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    Set toastValue = app.Toast("The quarterly report is ready to download now") _
        .MinWidth(100).MaxWidth(420)
    toastName = "rdm_wid72_" & toastValue.ComponentId
    With host.Shapes(toastName)
        transcript = transcript & "|fitsWide=" & CStr(.Width > 240 And .Width <= 420 _
            And .Height = 40)
    End With
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce

    ' A message cut at the default cap gets its words back when a higher
    ' MaxHeight follows, since a toast draws as soon as it is made.
    Set toastValue = app.Toast(longWords & " " & longWords & " " & longWords).MaxHeight(400)
    toastName = "rdm_wid72_" & toastValue.ComponentId
    With host.Shapes(toastName)
        transcript = transcript & "|raisedCapRestores=" & CStr(.Height > 160 _
            And Right$(.TextFrame2.TextRange.Text, 1) <> ChrW(8230))
    End With
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce
    ' A long path with no space to cut at still ends in an ellipsis that
    ' fits the capped card, by the card's own measure.
    Set toastValue = app.Toast("C:\Reports\" & String$(300, "q") & ".xlsx").MaxHeight(70)
    toastName = "rdm_wid72_" & toastValue.ComponentId
    With host.Shapes(toastName)
        With .TextFrame2
            needTall = .TextRange.BoundHeight + .MarginTop + .MarginBottom + 2.35 * 2 + 6
        End With
        transcript = transcript & "|tokenFits=" & CStr(.Height <= 70.01 _
            And Right$(.TextFrame2.TextRange.Text, 1) = ChrW(8230) _
            And needTall <= .Height + 0.5)
    End With
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce

    ' A tray placed near the window's bottom for a one-line toast lifts
    ' when its top toast runs taller, so the toast stays in the window.
    Set edgeHost = NewCanvas()
    Set edgeApp = ReDimUI.Mount(edgeHost, "wid74")
    viewBottom = ActiveWindow.VisibleRange.Top + ActiveWindow.VisibleRange.Height
    viewRight = ActiveWindow.VisibleRange.Left + ActiveWindow.VisibleRange.Width
    edgeApp.Label("low").AtRect(24, viewBottom - 30, 200, 20).Text "Low"
    edgeApp.Render
    Set toastValue = edgeApp.Toast(longWords)
    toastName = "rdm_wid74_" & toastValue.ComponentId
    With edgeHost.Shapes(toastName)
        transcript = transcript & "|tallLifts=" & CStr(.Height > 50 _
            And .Top + .Height <= viewBottom - 7.5)
    End With
    ReDimUI.DispatchShape toastName & "__tx"
    ReDimUI.PumpOnce

    ' A toast shown while its sheet is behind another fits again once the
    ' sheet is in front: as wide as the same toast shown there.
    edgeApp.Label("edge").AtRect(viewRight - 320, 24, 120, 20).Text "Edge"
    NewCanvas
    Set toastValue = edgeApp.Toast("The quarterly report is ready to download " & _
        "from the shared folder now").MinWidth(100).MaxWidth(420)
    toastName = "rdm_wid74_" & toastValue.ComponentId
    edgeHost.Activate
    ReDimUI.PumpOnce
    Set toastValue = edgeApp.Toast("The quarterly report is ready to download " & _
        "from the shared folder now").MinWidth(100).MaxWidth(420)
    transcript = transcript & "|offSheetRefits=" & CStr(Abs(edgeHost.Shapes(toastName).Width _
        - edgeHost.Shapes("rdm_wid74_" & toastValue.ComponentId).Width) < 0.5)
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestToastSize = transcript
End Function

' Keyboard focus: Tab order (TabIndex first, then creation order, -1
' skipped), the focus ring, Space and Enter per kind, Esc to leave, clicks
' that focus only text fields, the default button, access keys, the
' click-away and selection watches, and scrolling focus into view.
Public Function TestKeyboardFocus() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim visibleArea As Range

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid34")
    app.Button("first").AtRect(24, 24, 90, 26).Text("First").TabIndex 1
    app.TextInput("name").AtRect(24, 60, 150, 22).WritesTo "who"
    app.Button("save").AtRect(24, 96, 90, 26).Text("Save") _
        .OnClick "TestReDimWidgets.RecordChange"
    app.Toggle("dark").AtRect(24, 132, 44, 22).WritesTo "darkMode"
    app.TickBox("agree").AtRect(24, 166, 160, 20).Text("Agree").AccessKey "a"
    app.Button("later").AtRect(140, 24, 90, 26).Text("Later").TabIndex (-1)
    app.Render

    app.FocusFirst
    transcript = "firstFocused=" & CStr(ReDimUI.FocusedComponentId = "first")
    transcript = transcript & "|ringOffByDefault=" & _
        CStr(Not ShapeExists(host, "rdm_wid34_first__fr"))
    app.FocusRings
    transcript = transcript & "|ringDrawn=" & _
        CStr(ShapeExists(host, "rdm_wid34_first__fr"))
    app.FocusRings False
    transcript = transcript & "|ringOffTakesRing=" & _
        CStr(Not ShapeExists(host, "rdm_wid34_first__fr"))
    app.FocusRings
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabToField=" & _
        CStr(ReDimUI.FocusedComponentId = "name" And _
            Not ShapeExists(host, "rdm_wid34_first__fr"))
    RdxKeyChar "B"
    RdxKeyChar "o"
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabCommits=" & app.State("who")
    transcript = transcript & "|onButton=" & CStr(ReDimUI.FocusedComponentId = "save")
    RdxKeyChar " "
    transcript = transcript & "|spaceClicks=" & CStr(gChangeCount = 1)
    RdxKeyChar "{TAB}"
    RdxKeyChar " "
    transcript = transcript & "|spaceToggles=" & CStr(app.State("darkMode") = True)
    RdxKeyChar "{TAB}"
    RdxKeyChar " "
    transcript = transcript & "|spaceChecks=" & CStr(app.TickBox("agree").IsChecked)
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabWraps=" & CStr(ReDimUI.FocusedComponentId = "first")
    RdxKeyChar "{BACKTAB}"
    transcript = transcript & "|backTab=" & CStr(ReDimUI.FocusedComponentId = "agree")

    ' TabIndex -1 leaves a control to Focus alone.
    app.Button("later").Focus
    transcript = transcript & "|focusApi=" & CStr(ReDimUI.FocusedComponentId = "later")
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escLeaves=" & _
        CStr(Not ReDimUI.HasKeyboardFocus And _
            Not ShapeExists(host, "rdm_wid34_later__fr"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid34_save"
    transcript = transcript & "|clickNoFocus=" & _
        CStr(Not ReDimUI.HasKeyboardFocus And gChangeCount = 2)

    ' Enter in a text field commits and clicks the default button.
    app.DefaultButton "save"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid34_name"
    RdxKeyChar "x"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterDefault=" & _
        CStr(app.State("who") = "Box" And gChangeCount = 3)

    ' Alt plus the access key clicks its control; the key is underlined.
    transcript = transcript & "|keyUnderlined=" & _
        CStr(host.Shapes("rdm_wid34_agree__lbl").TextFrame2.TextRange _
            .Characters(1, 1).Font.UnderlineStyle = msoUnderlineSingleLine)
    ReDimUI.DispatchAccessKey "a"
    transcript = transcript & "|accessKeyClicks=" & _
        CStr(Not app.TickBox("agree").IsChecked)

    ' A moved selection and a click elsewhere both end focus.
    app.Button("save").Focus
    host.Range("E9").Select
    ReDimUI.PumpOnce
    transcript = transcript & "|selectionEndsFocus=" & CStr(Not ReDimUI.HasKeyboardFocus)
    app.Button("save").Focus
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid34_dark"
    transcript = transcript & "|clickEndsFocus=" & CStr(Not ReDimUI.HasKeyboardFocus)

    ' Focus below the visible window scrolls it into view.
    app.Button("far").AtRect(24, 2400, 90, 26).Text "Far"
    app.Render
    ActiveWindow.ScrollRow = 1
    app.Button("far").Focus
    Set visibleArea = ActiveWindow.VisibleRange
    transcript = transcript & "|scrolledIntoView=" & _
        CStr(Not Intersect(visibleArea, _
            host.Shapes("rdm_wid34_far").TopLeftCell) Is Nothing)
    RdxKeyChar "{ESC}"
    ActiveWindow.ScrollRow = 1
    app.Unmount False
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestKeyboardFocus = transcript
End Function

' Keys per kind: radio arrows move the selection and wrap, stepper and
' slider keys step and jump, check-list and transfer-list cursors, and the
' select's closed arrows, type-ahead, open list, and Esc.
Public Function TestControlKeys() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid35")
    app.RadioGroup("size").AtRect(24, 24, 120, 60) _
        .Items("Small", "Medium", "Large").Value(1).WritesTo "sizeState"
    app.RadioGroup("size").OnChange "TestReDimWidgets.RecordChange"
    app.Stepper("qty").AtRect(24, 100, 120, 24).SliderRange(0, 20, 1).Value 5
    app.SlideBar("vol").AtRect(24, 140, 160, 18).SliderRange(0, 100, 5).Value 50
    app.CheckList("feat").AtRect(220, 24, 170, 80).Items "Alpha", "Bravo", "Charlie"
    app.TransferList("pool").AtRect(220, 140, 360, 140).Items "One", "Two", "Three"
    app.SelectBox("zone").AtRect(24, 200, 150, 22) _
        .Items("North", "South", "East", "West", "Northwest").Value 1
    app.Render

    app.RadioGroup("size").Focus
    RdxKeyChar "{DOWN}"
    transcript = "radioDown=" & app.State("sizeState")
    RdxKeyChar "{UP}"
    RdxKeyChar "{UP}"
    transcript = transcript & "|radioWraps=" & app.State("sizeState")
    transcript = transcript & "|radioFires=" & gChangeCount
    RdxKeyChar "m"
    transcript = transcript & "|radioLetter=" & app.State("sizeState")
    RdxKeyChar "{PGUP}"
    transcript = transcript & "|radioPageUp=" & app.State("sizeState")
    RdxKeyChar "{PGDN}"
    transcript = transcript & "|radioPageDown=" & app.State("sizeState")

    RdxKeyChar "{TAB}"
    RdxKeyChar "{UP}"
    transcript = transcript & "|stepUp=" & CStr(app.Stepper("qty").CurrentValue)
    RdxKeyChar "{PGUP}"
    transcript = transcript & "|stepPage=" & CStr(app.Stepper("qty").CurrentValue)
    RdxKeyChar "{END}"
    transcript = transcript & "|stepEnd=" & CStr(app.Stepper("qty").CurrentValue)
    RdxKeyChar "{HOME}"
    transcript = transcript & "|stepHome=" & CStr(app.Stepper("qty").CurrentValue)

    RdxKeyChar "{TAB}"
    RdxKeyChar "{RIGHT}"
    RdxKeyChar "{PGDN}"
    transcript = transcript & "|slider=" & CStr(app.SlideBar("vol").CurrentValue)
    RdxKeyChar "{END}"
    transcript = transcript & "|sliderEnd=" & CStr(app.SlideBar("vol").CurrentValue)
    RdxKeyChar "3"
    RdxKeyChar "5"
    transcript = transcript & "|sliderTyped=" & CStr(app.SlideBar("vol").CurrentValue)

    RdxKeyChar "{TAB}"
    transcript = transcript & "|cursorDrawn=" & _
        CStr(ShapeExists(host, "rdm_wid35_feat__kr"))
    RdxKeyChar " "
    RdxKeyChar "{DOWN}"
    RdxKeyChar " "
    transcript = transcript & "|checkKeys=" & _
        CStr(app.CheckList("feat").CheckedCount = 2 And _
            app.CheckList("feat").IsItemChecked(1) And _
            app.CheckList("feat").IsItemChecked(2))
    RdxKeyChar "{HOME}"
    RdxKeyChar " "
    transcript = transcript & "|headerKey=" & _
        CStr(app.CheckList("feat").CheckedCount = 3)

    RdxKeyChar "{TAB}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterMoves=" & _
        CStr(app.TransferList("pool").ChosenCount = 1 And _
            app.TransferList("pool").ChosenTextAt(1) = "Two")
    RdxKeyChar "{RIGHT}"
    RdxKeyChar " "
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|movesBack=" & _
        CStr(app.TransferList("pool").ChosenCount = 0)

    RdxKeyChar "{TAB}"
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|selectArrow=" & CStr(app.SelectBox("zone").CurrentValue)
    RdxKeyChar "n"
    transcript = transcript & "|typeAhead=" & CStr(app.SelectBox("zone").CurrentValue)
    RdxKeyChar " "
    transcript = transcript & "|spaceOpens=" & _
        CStr(host.Shapes("rdm_wid35_zone__opt5").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)
    RdxKeyChar "{HOME}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|openPick=" & _
        CStr(app.SelectBox("zone").CurrentValue = 2 And _
            Not ShapeExists(host, "rdm_wid35_zone__opt1"))
    RdxKeyChar "{F4}"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escCloses=" & _
        CStr(Not ShapeExists(host, "rdm_wid35_zone__opt1") And _
            ReDimUI.HasKeyboardFocus)
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escLeaves=" & CStr(Not ReDimUI.HasKeyboardFocus)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestControlKeys = transcript
End Function

' The modal takes focus on OK, Tab stays inside it, Enter confirms and
' returns focus to where it was, and Esc cancels.
Public Function TestModalKeys() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gConfirmRan = 0
    gCancelRan = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid36")
    app.Button("del").AtRect(24, 24, 90, 26).Text "Delete"
    app.TextInput("note").AtRect 24, 60, 150, 22
    app.Render
    app.Button("del").Focus
    app.Confirm "Delete rows", "Remove 42 rows?", _
        "TestReDimWidgets.RecordConfirm", "TestReDimWidgets.RecordCancelChoice"
    transcript = "modalTakesFocus=" & CStr(ReDimUI.FocusedComponentId = "mdl_ok")
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabToCancel=" & _
        CStr(ReDimUI.FocusedComponentId = "mdl_cancel")
    RdxKeyChar "{TAB}"
    transcript = transcript & "|trapped=" & CStr(ReDimUI.FocusedComponentId = "mdl_ok")
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterConfirms=" & CStr(gConfirmRan = 1)
    transcript = transcript & "|focusReturns=" & CStr(ReDimUI.FocusedComponentId = "del")
    app.Confirm "Again", "Sure?", _
        "TestReDimWidgets.RecordConfirm", "TestReDimWidgets.RecordCancelChoice"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escCancels=" & _
        CStr(gCancelRan = 1 And gConfirmRan = 1)
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestModalKeys = transcript
End Function

' Clearable: the button shows only while the field holds text, empties it
' and keeps focus, and a combo's list opens unfiltered.
Public Function TestClearable() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid37")
    app.TextInput("q").AtRect(24, 24, 160, 22).Clearable.WritesTo "query"
    app.ComboBox("c").AtRect(24, 60, 160, 22).Items("Red", "Green").Clearable
    app.Render
    transcript = "hiddenWhenEmpty=" & CStr(Not ShapeExists(host, "rdm_wid37_q__cx"))
    ReDimUI.DispatchShape "rdm_wid37_q"
    RdxKeyChar "a"
    RdxKeyChar "b"
    transcript = transcript & "|shownWithText=" & _
        CStr(ShapeExists(host, "rdm_wid37_q__cx"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid37_q__cx"
    transcript = transcript & "|clearEmpties=" & _
        CStr(app.TextInput("q").InputValue = vbNullString And _
            ReDimUI.FocusedComponentId = "q" And _
            Not ShapeExists(host, "rdm_wid37_q__cx"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid37_c"
    RdxKeyChar "r"
    RdxKeyChar "e"
    transcript = transcript & "|comboX=" & CStr(ShapeExists(host, "rdm_wid37_c__cx"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid37_c__cx"
    transcript = transcript & "|comboCleared=" & _
        CStr(app.ComboBox("c").InputValue = vbNullString And _
            ShapeExists(host, "rdm_wid37_c__opt2"))
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestClearable = transcript
End Function

' Text editing in float fields: Shift selection painted as a highlight,
' the clipboard through Ctrl+C, Ctrl+X, and Ctrl+V, undo and redo with
' typing grouped, Ctrl+A, word moves and deletes, and symbol keys.
Public Function TestTextEditing() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim transcript As String
    Dim keyNo As Long

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid38")
    app.TextInput("one").AtRect 24, 24, 220, 22
    app.TextInput("two").AtRect 24, 60, 220, 22
    app.Render
    Set fieldShape = host.Shapes("rdm_wid38_one")

    ReDimUI.DispatchShape "rdm_wid38_one"
    For keyNo = 1 To Len("hello world")
        RdxKeyChar Mid$("hello world", keyNo, 1)
    Next keyNo
    For keyNo = 1 To 5
        RdxKeyChar "{SHIFTLEFT}"
    Next keyNo
    ' The face reads "hello |world": the bar sits before the selection.
    transcript = "selectionPainted=" & _
        CStr(fieldShape.TextFrame2.TextRange.Characters(8, 1).Font.Highlight.RGB _
            = app.Theme.PrimaryColor And _
            fieldShape.TextFrame2.TextRange.Characters(1, 1).Font.Highlight.RGB _
            <> app.Theme.PrimaryColor)
    RdxKeyChar "{COPY}"
    RdxKeyChar "{CUT}"
    transcript = transcript & "|cut=" & app.TextInput("one").InputValue
    RdxKeyChar "{UNDO}"
    transcript = transcript & "|undo=" & app.TextInput("one").InputValue
    RdxKeyChar "{REDO}"
    transcript = transcript & "|redo=" & app.TextInput("one").InputValue

    ' Paste into the other field, replacing its selection.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid38_two"
    RdxKeyChar "a"
    RdxKeyChar "b"
    RdxKeyChar "{SELECTALL}"
    RdxKeyChar "{PASTE}"
    transcript = transcript & "|pasted=" & app.TextInput("two").InputValue
    RdxKeyChar "{UNDO}"
    transcript = transcript & "|pasteUndone=" & app.TextInput("two").InputValue
    RdxKeyChar "{UNDO}"
    transcript = transcript & "|typingOneStep=" & _
        CStr(LenB(app.TextInput("two").InputValue) = 0)

    ' Word moves and deletes.
    RdxKeyChar "{SELECTALL}"
    For keyNo = 1 To Len("one two three")
        RdxKeyChar Mid$("one two three", keyNo, 1)
    Next keyNo
    RdxKeyChar "{WORDLEFT}"
    RdxKeyChar "_"
    transcript = transcript & "|wordLeft=" & app.TextInput("two").InputValue
    RdxKeyChar "{TEXTEND}"
    RdxKeyChar "{WORDBS}"
    transcript = transcript & "|wordBackspace=" & app.TextInput("two").InputValue
    RdxKeyChar "{TEXTHOME}"
    RdxKeyChar "{SHIFTWORDRIGHT}"
    RdxKeyChar "{DEL}"
    transcript = transcript & "|shiftWordDelete=" & app.TextInput("two").InputValue

    ' Symbols, the apostrophe and quote included.
    RdxKeyChar "{SELECTALL}"
    RdxKeyChar "@"
    RdxKeyChar "{APOS}"
    RdxKeyChar "{QUOTE}"
    RdxKeyChar "("
    RdxKeyChar "~"
    transcript = transcript & "|symbols=" & app.TextInput("two").InputValue
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestTextEditing = transcript
End Function

' Points the pointer override at a fraction across one face character.
Private Sub PointAtChar(ByVal fieldShape As Shape, ByVal charIndex As Long, ByVal across As Double)
    With fieldShape.TextFrame2.TextRange.Characters(charIndex, 1)
        ReDimUI.OverridePointer .BoundLeft + .BoundWidth * across, _
            .BoundTop + .BoundHeight / 2
    End With
End Sub

' Clicks in a float field: the first click focuses it and puts the caret
' at the character under the pointer, a later click moves the caret, and
' a double click selects the word under the pointer.
Public Function TestFieldClicks() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim faceIndex As Long
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid39")
    app.TextInput("f").AtRect 24, 24, 260, 22
    app.TextInput("f").InputValue = "alpha beta gamma"
    app.Render
    Set fieldShape = host.Shapes("rdm_wid39_f")

    ' The left half of the b in beta, character 7: the caret goes before it.
    PointAtChar fieldShape, 7, 0.25
    ReDimUI.DispatchShape "rdm_wid39_f"
    RdxKeyChar "X"
    transcript = "clickPlacesCaret=" & app.TextInput("f").InputValue

    ' The right half of the l in alpha, a later click: the caret follows.
    faceIndex = InStr(fieldShape.TextFrame2.TextRange.Text, "l")
    PointAtChar fieldShape, faceIndex, 0.75
    Sleep 600
    ReDimUI.DispatchShape "rdm_wid39_f"
    RdxKeyChar "Y"
    transcript = transcript & "|clickMovesCaret=" & app.TextInput("f").InputValue

    ' Two quick clicks on the m in gamma select the word.
    faceIndex = InStr(fieldShape.TextFrame2.TextRange.Text, "gamma") + 2
    PointAtChar fieldShape, faceIndex, 0.5
    Sleep 600
    ReDimUI.DispatchShape "rdm_wid39_f"
    ReDimUI.DispatchShape "rdm_wid39_f"
    RdxKeyChar "Z"
    transcript = transcript & "|doubleClickWord=" & app.TextInput("f").InputValue
    ReDimUI.ClearPointerOverride
    RdxKeyChar "{ENTER}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFieldClicks = transcript
End Function

Public Sub RecordInput()
    gInputCount = gInputCount + 1
    gLastInput = ReDimUI.Sender.InputValue
End Sub

Public Function CheckEmail(ByVal candidate As String) As String
    If InStr(candidate, "@") = 0 Then CheckEmail = "Needs an @"
End Function

Private Sub TypeText(ByVal typedText As String)
    Dim keyNo As Long

    For keyNo = 1 To Len(typedText)
        RdxKeyChar Mid$(typedText, keyNo, 1)
    Next keyNo
End Sub

' Field rules: a placeholder in muted ink, Numeric's filter, MaxLength
' with its count, Validates with its border and message rechecked live,
' and OnInput now or after a pause.
Public Function TestFieldRules() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim separatorChar As String

    gInputCount = 0
    gLastInput = vbNullString
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid40")
    app.TextInput("find").AtRect(24, 24, 200, 22).Placeholder "Search..."
    app.TextInput("qty").AtRect(24, 60, 120, 22).Numeric
    app.TextInput("code").AtRect(24, 96, 160, 22).MaxLength 5
    app.TextInput("mail").AtRect(24, 150, 200, 22).Validates "TestReDimWidgets.CheckEmail"
    app.TextInput("live").AtRect(260, 24, 160, 22).OnInput "TestReDimWidgets.RecordInput"
    app.Render

    transcript = "placeholderShown=" & _
        CStr(host.Shapes("rdm_wid40_find").TextFrame2.TextRange.Text = "Search..." _
            And InkOf(host, "rdm_wid40_find") = app.Theme.OnMutedColor)
    ReDimUI.DispatchShape "rdm_wid40_find"
    transcript = transcript & "|placeholderFocused=" & _
        CStr(host.Shapes("rdm_wid40_find").TextFrame2.TextRange.Text = "|Search...")
    RdxKeyChar "a"
    transcript = transcript & "|placeholderGone=" & _
        CStr(host.Shapes("rdm_wid40_find").TextFrame2.TextRange.Text = "a|" _
            And InkOf(host, "rdm_wid40_find") = app.Theme.OnSurfaceColor)

    ReDimUI.DispatchShape "rdm_wid40_qty"
    TypeText "1a2.3.4-"
    RdxKeyChar "{TEXTHOME}"
    RdxKeyChar "-"
    If Application.UseSystemSeparators Then
        separatorChar = Application.International(xlDecimalSeparator)
    Else
        separatorChar = Application.DecimalSeparator
    End If
    transcript = transcript & "|numeric=" & _
        CStr(app.TextInput("qty").InputValue = "-12" & separatorChar & "34")
    ' The keypad's decimal key types the separator Excel uses.
    app.TextInput("qty").InputValue = "7"
    RdxKeyChar "{TEXTEND}"
    RdxKeyChar "{DECIMAL}"
    RdxKeyChar "5"
    transcript = transcript & "|keypadDecimal=" & _
        CStr(app.TextInput("qty").InputValue = "7" & separatorChar & "5")

    ReDimUI.DispatchShape "rdm_wid40_code"
    TypeText "abcdefg"
    transcript = transcript & "|maxLength=" & app.TextInput("code").InputValue
    transcript = transcript & "|counter=" & _
        host.Shapes("rdm_wid40_code__mc").TextFrame2.TextRange.Text

    ReDimUI.DispatchShape "rdm_wid40_mail"
    TypeText "bob"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|invalid=" & app.TextInput("mail").ValidationError
    transcript = transcript & "|messageShown=" & _
        CStr(host.Shapes("rdm_wid40_mail__me").TextFrame2.TextRange.Text = "Needs an @")
    transcript = transcript & "|dangerBorder=" & _
        CStr(host.Shapes("rdm_wid40_mail").Line.ForeColor.RGB = app.Theme.DangerColor)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid40_mail"
    RdxKeyChar "@"
    transcript = transcript & "|liveRecheck=" & _
        CStr(LenB(app.TextInput("mail").ValidationError) = 0 And _
            Not ShapeExists(host, "rdm_wid40_mail__me"))
    RdxKeyChar "{ENTER}"

    ReDimUI.DispatchShape "rdm_wid40_live"
    TypeText "xy"
    transcript = transcript & "|inputNow=" & gInputCount & ":" & gLastInput
    app.TextInput("live").DebounceMs 150
    RdxKeyChar "z"
    transcript = transcript & "|debounceWaits=" & CStr(gInputCount = 2)
    Sleep 200
    ReDimUI.PumpOnce
    transcript = transcript & "|debounceFires=" & gInputCount & ":" & gLastInput
    RdxKeyChar "{ENTER}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFieldRules = transcript
End Function

' AutoGrow: a multi-line field grows a line at a time from the height it
' was given up to its cap and shrinks back as lines go, a component placed
' Below it moves with it, and a first render starts at the grown height.
Public Function TestAutoGrow() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim lineStep As Double
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid41")
    app.TextInput("notes").AtRect(24, 24, 220, 22).AutoGrow 3
    app.Label("after").Below("notes", 6).Sized(220, 18).Text "Below the notes"
    app.TextInput("pre").AtRect(280, 24, 220, 22).AutoGrow
    app.TextInput("pre").InputValue = "first" & vbLf & "second"
    app.Render
    Set fieldShape = host.Shapes("rdm_wid41_notes")
    lineStep = app.Theme.BaseFontSize * 1.35

    transcript = "startsAtGiven=" & CStr(Abs(fieldShape.Height - 22) < 0.01)
    transcript = transcript & "|firstRenderGrown=" & _
        CStr(Abs(host.Shapes("rdm_wid41_pre").Height - (2 * lineStep + 6)) < 0.01)

    ReDimUI.DispatchShape "rdm_wid41_notes"
    TypeText "one"
    RdxKeyChar "{ENTER}"
    TypeText "two"
    transcript = transcript & "|grewTwoLines=" & _
        CStr(Abs(fieldShape.Height - (2 * lineStep + 6)) < 0.01)
    transcript = transcript & "|belowFollows=" & _
        CStr(Abs(host.Shapes("rdm_wid41_after").Top _
            - (fieldShape.Top + fieldShape.Height + 6)) < 0.01)
    RdxKeyChar "{ENTER}"
    TypeText "three"
    RdxKeyChar "{ENTER}"
    TypeText "four"
    transcript = transcript & "|capped=" & _
        CStr(Abs(fieldShape.Height - (3 * lineStep + 6)) < 0.01)
    transcript = transcript & "|faceScrolls=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = _
            ChrW(8230) & "two" & vbLf & "three" & vbLf & "four|")

    RdxKeyChar "{SELECTALL}"
    RdxKeyChar "{BS}"
    transcript = transcript & "|shrinksBack=" & _
        CStr(Abs(fieldShape.Height - 22) < 0.01)
    transcript = transcript & "|belowReturns=" & _
        CStr(Abs(host.Shapes("rdm_wid41_after").Top - 52) < 0.01)

    ' At rest a long line wraps, and the field grows to hold it.
    RdxKeyChar "{CTRLENTER}"
    app.TextInput("notes").InputValue = String(90, "x")
    transcript = transcript & "|wrapGrows=" & CStr(fieldShape.Height > 23)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestAutoGrow = transcript
End Function

' The bold letters of a shape's text, in order.
Private Function BoldRun(ByVal host As Worksheet, ByVal shapeName As String) As String
    Dim charNo As Long

    With host.Shapes(shapeName).TextFrame2.TextRange
        For charNo = 1 To .Length
            If .Characters(charNo, 1).Font.Bold = msoTrue Then
                BoldRun = BoldRun & .Characters(charNo, 1).Text
            End If
        Next charNo
    End With
End Function

' Combo assists: rows bold what the text matched; the first item the text
' begins shows its rest after the caret in muted ink, which Right at the
' end or Tab takes and Undo gives back; RestrictToItems commits only
' items, a prefix, or an empty text, and otherwise reverts.
Public Function TestComboAssists() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim fieldShape As Shape
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid42")
    app.ComboBox("fruit").AtRect 24, 24, 200, 22
    app.ComboBox("fruit").Items "Apple", "Banana", "Blueberry", "Cherry"
    app.ComboBox("strict").AtRect 260, 24, 200, 22
    app.ComboBox("strict").Items("Apple", "Banana", "Blueberry", "Cherry") _
        .RestrictToItems
    app.Render
    Set fieldShape = host.Shapes("rdm_wid42_fruit")

    ReDimUI.DispatchShape "rdm_wid42_fruit"
    RdxKeyChar "b"
    transcript = "boldMatch=" & BoldRun(host, "rdm_wid42_fruit__opt1") & "," & _
        BoldRun(host, "rdm_wid42_fruit__opt2")
    transcript = transcript & "|ghostShown=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "b|anana")
    transcript = transcript & "|ghostMuted=" & _
        CStr(fieldShape.TextFrame2.TextRange.Characters(3, 5).Font.Fill.ForeColor.RGB _
            = app.Theme.OnMutedColor _
        And fieldShape.TextFrame2.TextRange.Characters(1, 1).Font.Fill.ForeColor.RGB _
            = app.Theme.OnSurfaceColor)
    RdxKeyChar "a"
    transcript = transcript & "|boldGrows=" & _
        BoldRun(host, "rdm_wid42_fruit__opt1")
    RdxKeyChar "{BS}"
    transcript = transcript & "|boldShrinks=" & _
        BoldRun(host, "rdm_wid42_fruit__opt1")
    RdxKeyChar "{BS}"
    transcript = transcript & "|boldCleared=" & _
        CStr(LenB(BoldRun(host, "rdm_wid42_fruit__opt1")) = 0 _
            And LenB(BoldRun(host, "rdm_wid42_fruit__opt2")) = 0)
    TypeText "bl"
    transcript = transcript & "|ghostFollows=" & _
        CStr(fieldShape.TextFrame2.TextRange.Text = "bl|ueberry") & _
        ":" & BoldRun(host, "rdm_wid42_fruit__opt1")
    RdxKeyChar "{RIGHT}"
    transcript = transcript & "|rightTakes=" & app.ComboBox("fruit").InputValue
    RdxKeyChar "{UNDO}"
    transcript = transcript & "|undoGivesBack=" & app.ComboBox("fruit").InputValue
    BackspaceAll app.ComboBox("fruit")
    RdxKeyChar "c"
    RdxKeyChar "{TAB}"
    transcript = transcript & "|tabTakes=" & app.ComboBox("fruit").InputValue & _
        ":" & CStr(ReDimUI.FocusedComponentId = "strict")
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid42_strict"
    TypeText "zz"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|restrictReverts=" & _
        CStr(LenB(app.ComboBox("strict").InputValue) = 0)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid42_strict"
    TypeText "blu"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|restrictCompletes=" & app.ComboBox("strict").InputValue
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid42_strict"
    BackspaceAll app.ComboBox("strict")
    TypeText "APPLE"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|restrictSpelling=" & app.ComboBox("strict").InputValue
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid42_strict"
    BackspaceAll app.ComboBox("strict")
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|restrictEmpty=" & _
        CStr(LenB(app.ComboBox("strict").InputValue) = 0)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestComboAssists = transcript
End Function

' The pointer seam drives the pump's press watch: a held button on a
' slider drags it with the value in a bubble over the thumb, the release
' fires OnChange once, keyboard focus shows the bubble too, and a toast
' under the pointer holds its countdown until the pointer leaves.
Public Function TestPointerBasics() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim toastValue As ReDimUI
    Dim toastName As String
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    ReDimUI.ReduceMotion True
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid43")
    app.SlideBar("vol").AtRect(24, 60, 200, 18).SliderRange(0, 100, 10).Value(20) _
        .OnChange "TestReDimWidgets.RecordChange"
    app.Render

    ReDimUI.OverridePointer 124, 69, True
    ReDimUI.PumpOnce
    transcript = "pressDrags=" & CStr(app.SlideBar("vol").IsSlideDragging) & _
        ":" & app.SlideBar("vol").CurrentValue
    transcript = transcript & "|bubbleShown=" & _
        host.Shapes("rdm_wid43_vol__vb").TextFrame2.TextRange.Text
    transcript = transcript & "|bubbleAbove=" & _
        CStr(host.Shapes("rdm_wid43_vol__vb").Top < 60)
    ReDimUI.OverridePointer 184, 69, True
    ReDimUI.PumpOnce
    transcript = transcript & "|bubbleFollows=" & _
        host.Shapes("rdm_wid43_vol__vb").TextFrame2.TextRange.Text
    ReDimUI.OverridePointer 184, 69, False
    ReDimUI.PumpOnce
    transcript = transcript & "|releaseCommits=" & CStr(gChangeCount) & ":" & _
        app.SlideBar("vol").CurrentValue
    transcript = transcript & "|bubbleGone=" & _
        CStr(Not ShapeExists(host, "rdm_wid43_vol__vb"))

    app.SlideBar("vol").Focus
    RdxKeyChar "{RIGHT}"
    transcript = transcript & "|focusBubble=" & _
        host.Shapes("rdm_wid43_vol__vb").TextFrame2.TextRange.Text
    RdxKeyChar "{ESC}"
    transcript = transcript & "|blurHides=" & _
        CStr(Not ShapeExists(host, "rdm_wid43_vol__vb"))

    Set toastValue = app.Toast("Hold me.", 300)
    toastName = "rdm_wid43_" & toastValue.ComponentId
    With host.Shapes(toastName)
        ReDimUI.OverridePointer .Left + .Width / 2, .Top + .Height / 2
    End With
    Sleep 400
    ReDimUI.PumpOnce
    transcript = transcript & "|toastHeld=" & _
        CStr(ShapeExists(host, toastName) And toastValue.ToastRemainingMs >= 900)
    ReDimUI.OverridePointer 1, 1
    Sleep 1100
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|toastResumes=" & CStr(Not ShapeExists(host, toastName))

    ReDimUI.ClearPointerOverride
    RdxReleaseKeys
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestPointerBasics = transcript
End Function

Private Function FillOf(ByVal host As Worksheet, ByVal shapeName As String) As Long
    FillOf = host.Shapes(shapeName).Fill.ForeColor.RGB
End Function

' Pointer effects: the control under the pointer takes a hover tint and a
' stronger one while the button that went down on it stays down; parts
' answer for themselves (a stepper's plus, a check-list row's box, a
' transfer row and move button); an open list's highlight follows a
' moving pointer; leaving clears the look; and a sheet coming back to the
' front re-arms the pump the effects need.
Public Function TestPointerEffects() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim baseFill As Long
    Dim hoverFill As Long
    Dim eventsWereOn As Boolean
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid44")
    app.PointerEffects
    app.Button("go").AtRect(24, 24, 100, 30).Text "Go"
    app.Stepper("qty").AtRect(24, 70, 120, 24).SliderRange(0, 9, 1).Value 3
    app.CheckList("chk").AtRect(24, 110, 160, 80).Items "A", "B", "C"
    app.TransferList("tl").AtRect(220, 24, 300, 160).Items "One", "Two", "Three"
    app.SelectBox("sel").AtRect(24, 220, 140, 22).Items "Red", "Green", "Blue"
    app.Render
    baseFill = FillOf(host, "rdm_wid44_go")

    ReDimUI.OverridePointer 74, 39
    ReDimUI.PumpOnce
    hoverFill = FillOf(host, "rdm_wid44_go")
    transcript = "buttonHover=" & CStr(hoverFill <> baseFill)
    ReDimUI.OverridePointer 74, 39, True
    ReDimUI.PumpOnce
    transcript = transcript & "|buttonPressed=" & _
        CStr(FillOf(host, "rdm_wid44_go") <> hoverFill _
            And FillOf(host, "rdm_wid44_go") <> baseFill)
    ReDimUI.OverridePointer 74, 39, False
    ReDimUI.PumpOnce
    transcript = transcript & "|releaseHover=" & _
        CStr(FillOf(host, "rdm_wid44_go") = hoverFill)

    ' A press that went down elsewhere does not press what it crosses.
    ReDimUI.OverridePointer 600, 300, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 74, 39, True
    ReDimUI.PumpOnce
    transcript = transcript & "|crossingNotPressed=" & _
        CStr(FillOf(host, "rdm_wid44_go") = hoverFill)
    ReDimUI.OverridePointer 74, 39, False
    ReDimUI.PumpOnce

    ReDimUI.OverridePointer 132, 82
    ReDimUI.PumpOnce
    transcript = transcript & "|stepperPart=" & _
        CStr(FillOf(host, "rdm_wid44_qty__plus") <> FillOf(host, "rdm_wid44_qty__minus"))
    transcript = transcript & "|buttonCleared=" & _
        CStr(FillOf(host, "rdm_wid44_go") = baseFill)

    ReDimUI.OverridePointer 30, 160
    ReDimUI.PumpOnce
    transcript = transcript & "|checkRow=" & _
        CStr(host.Shapes("rdm_wid44_chk__b2").Line.ForeColor.RGB = app.Theme.PrimaryColor _
            And host.Shapes("rdm_wid44_chk").Line.ForeColor.RGB = app.Theme.BorderColor)

    ReDimUI.OverridePointer 230, 75
    ReDimUI.PumpOnce
    transcript = transcript & "|transferRow=" & _
        CStr(FillOf(host, "rdm_wid44_tl__al2") <> FillOf(host, "rdm_wid44_tl__al1"))
    ReDimUI.OverridePointer 370, 70
    ReDimUI.PumpOnce
    transcript = transcript & "|transferButton=" & _
        CStr(FillOf(host, "rdm_wid44_tl__mvr") <> FillOf(host, "rdm_wid44_tl__mvar"))
    transcript = transcript & "|rowCleared=" & _
        CStr(FillOf(host, "rdm_wid44_tl__al2") = FillOf(host, "rdm_wid44_tl__al1"))

    ReDimUI.DispatchShape "rdm_wid44_sel"
    With host.Shapes("rdm_wid44_sel__opt2")
        ReDimUI.OverridePointer .Left + 10, .Top + .Height / 2
    End With
    ReDimUI.PumpOnce
    transcript = transcript & "|listFollows=" & _
        CStr(FillOf(host, "rdm_wid44_sel__opt2") = app.Theme.PrimaryColor)
    ReDimUI.DispatchShape "rdm_wid44_sel"

    ReDimUI.OverridePointer 74, 39
    ReDimUI.PumpOnce
    app.PointerEffects False
    ReDimUI.PumpOnce
    transcript = transcript & "|effectsOff=" & _
        CStr(FillOf(host, "rdm_wid44_go") = baseFill)

    ' Leaving the sheet stops the pump's reason to run; coming back
    ' re-arms it. The harness may hold EnableEvents off, so the step
    ' pins it on and restores it.
    app.PointerEffects True
    ReDimUI.AutoPump True
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = True
    ActiveWorkbook.Worksheets.Add
    RdxStopPump
    host.Activate
    transcript = transcript & "|rearmed=" & CStr(RdxPumpArmed())
    Application.EnableEvents = eventsWereOn
    ' Tear down with the pump off: a real timer left armed would fire
    ' between the harness's runs.
    ReDimUI.AutoPump False
    RdxStopPump
    app.PointerEffects False
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestPointerEffects = transcript
End Function

' Tooltips: once the pointer rests on a control for the double-click time
' its tooltip shows, and a pointer still moving starts the wait over. A
' field-sized control's tip sits clear below it; a tall control's sits
' under the pointer and goes when the pointer reaches it. A press puts a
' tip away until the pointer leaves; a disabled control shows its
' DisabledReason; leaving hides it; an open list shows none; a tip fits
' its words and wraps a long one; and the alternative text carries both.
Public Function TestTooltips() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim restMs As Long
    Dim transcript As String
    Dim tipX As Double
    Dim tipY As Double
    Dim viewBottom As Double

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid45")
    app.Button("save").AtRect(24, 24, 100, 30).Text("Save").Tooltip "Saves the form"
    app.Button("send").AtRect(24, 70, 100, 30).Text("Send").Enabled(False) _
        .DisabledReason "Fill in the address first"
    app.Button("tall").AtRect(300, 24, 140, 160).Text("Tall").Tooltip "A tall control"
    app.SelectBox("pick").AtRect(500, 24, 120, 22).Items("A", "B", "C") _
        .Tooltip "Pick a letter"
    app.Button("long").AtRect(24, 240, 100, 30).Text("Long").Tooltip _
        "A long tooltip wraps once it runs past the widest a tip may grow, " & _
        "so its lines stay short enough to read at a glance."
    app.Render
    restMs = GetDoubleClickTime() + 60

    transcript = "altTip=" & host.Shapes("rdm_wid45_save").AlternativeText
    transcript = transcript & "|altReason=" & host.Shapes("rdm_wid45_send").AlternativeText
    ReDimUI.OverridePointer 40, 39
    ReDimUI.PumpOnce
    transcript = transcript & "|waits=" & CStr(Not ShapeExists(host, "rdm_wid45_save__tt"))
    Sleep restMs \ 2
    ReDimUI.OverridePointer 74, 39
    ReDimUI.PumpOnce
    Sleep restMs \ 2 + 30
    ReDimUI.PumpOnce
    transcript = transcript & "|movingWaits=" & _
        CStr(Not ShapeExists(host, "rdm_wid45_save__tt"))
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|tipShows=" & _
        host.Shapes("rdm_wid45_save__tt").TextFrame2.TextRange.Text
    With host.Shapes("rdm_wid45_save__tt")
        transcript = transcript & "|clearsControl=" & CStr(.Top >= 54 _
            And .Width < 150 And .Width > .TextFrame2.TextRange.BoundWidth)
    End With

    ReDimUI.OverridePointer 74, 39, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 74, 39, False
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|pressHides=" & CStr(Not ShapeExists(host, "rdm_wid45_save__tt"))

    ReDimUI.OverridePointer 74, 85
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|reasonShows=" & _
        host.Shapes("rdm_wid45_send__tt").TextFrame2.TextRange.Text
    ReDimUI.OverridePointer 600, 400
    ReDimUI.PumpOnce
    transcript = transcript & "|leaveHides=" & CStr(Not ShapeExists(host, "rdm_wid45_send__tt"))

    ' A tall control's tip sits under the pointer, inside the control, and
    ' goes when the pointer reaches it, so a click there lands on the control.
    ReDimUI.OverridePointer 320, 40
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    With host.Shapes("rdm_wid45_tall__tt")
        transcript = transcript & "|tallUnderPointer=" & CStr(.Top >= 40 And .Top < 184)
        tipX = .Left + .Width / 2
        tipY = .Top + .Height / 2
    End With
    ReDimUI.OverridePointer tipX, tipY
    ReDimUI.PumpOnce
    transcript = transcript & "|reachedHides=" & CStr(Not ShapeExists(host, "rdm_wid45_tall__tt"))
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|staysAway=" & CStr(Not ShapeExists(host, "rdm_wid45_tall__tt"))

    ' An open list shows no tip over its rows.
    ReDimUI.DispatchShape "rdm_wid45_pick"
    ReDimUI.OverridePointer 540, 58
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|openListNoTip=" & CStr(ShapeExists(host, "rdm_wid45_pick__opt1") _
        And Not ShapeExists(host, "rdm_wid45_pick__tt"))
    ReDimUI.DispatchShape "rdm_wid45_pick"

    ' A long tip wraps at its widest.
    ReDimUI.OverridePointer 60, 255
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    With host.Shapes("rdm_wid45_long__tt")
        transcript = transcript & "|longWraps=" & CStr(.Width <= 241 And .Height > 30)
    End With
    ' A tip whose words change while it shows, as when its control is
    ' disabled under the pointer, shows the new words.
    app.Button("long").DisabledReason("Not yet").Enabled False
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|tipFollowsWords=" & _
        host.Shapes("rdm_wid45_long__tt").TextFrame2.TextRange.Text
    app.Button("long").Enabled True

    ' With no room below the pointer, a tall control's tip sits above it
    ' as far clear as it would below, so a small move up keeps it.
    viewBottom = ActiveWindow.VisibleRange.Top + ActiveWindow.VisibleRange.Height
    app.Button("low").AtRect(150, viewBottom - 110, 100, 100).Text("Low") _
        .Tooltip "Near the bottom"
    ReDimUI.OverridePointer 200, 200
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 200, viewBottom - 20
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    With host.Shapes("rdm_wid45_low__tt")
        transcript = transcript & "|flipClear=" & CStr(.Top + .Height <= viewBottom - 36)
    End With
    ReDimUI.OverridePointer 200, viewBottom - 24
    ReDimUI.PumpOnce
    transcript = transcript & "|flipHolds=" & CStr(ShapeExists(host, "rdm_wid45_low__tt"))

    ' A label's tip goes with the label, and a tip showing at shutdown goes.
    app.Label("note").AtRect(460, 240, 120, 20).Text("Note").Tooltip "A note"
    ReDimUI.OverridePointer 200, 200
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 500, 250
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|labelTip=" & CStr(ShapeExists(host, "rdm_wid45_note__tt"))
    app.Component("note").Remove
    transcript = transcript & "|removeTakesTip=" & _
        CStr(Not ShapeExists(host, "rdm_wid45_note__tt"))
    ReDimUI.OverridePointer 74, 39
    ReDimUI.PumpOnce
    Sleep restMs
    ReDimUI.PumpOnce
    transcript = transcript & "|tipBeforeShutdown=" & _
        CStr(ShapeExists(host, "rdm_wid45_save__tt"))
    ReDimUI.Shutdown
    transcript = transcript & "|shutdownTakesTip=" & _
        CStr(Not ShapeExists(host, "rdm_wid45_save__tt"))
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestTooltips = transcript
End Function

' Hold-to-repeat: a press held on a stepper's plus steps after the
' keyboard repeat delay and keeps stepping at the repeat rate, writing its
' state as it goes; the release fires OnChange once and swallows the click
' Excel delivers for it. A short press repeats nothing, and its click
' steps once. A held transfer paging arrow pages again and again. The
' waits cover the slowest Windows settings: a 1000 ms delay, 400 ms rate.
Public Function TestHoldRepeat() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim itemNo As Long
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid46")
    app.Stepper("qty").AtRect(24, 24, 120, 24).SliderRange(0, 50, 1).Value(5) _
        .WritesTo("qty").OnChange "TestReDimWidgets.RecordChange"
    app.TransferList("tl").AtRect 220, 24, 300, 160
    For itemNo = 1 To 30
        app.TransferList("tl").AddItem "Item " & itemNo
    Next itemNo
    app.Render

    ReDimUI.OverridePointer 132, 36, True
    ReDimUI.PumpOnce
    transcript = "noStepAtPress=" & app.Stepper("qty").CurrentValue
    Sleep 1100
    ReDimUI.PumpOnce
    transcript = transcript & "|firstRepeat=" & app.Stepper("qty").CurrentValue
    Sleep 450
    ReDimUI.PumpOnce
    transcript = transcript & "|keepsRepeating=" & app.Stepper("qty").CurrentValue
    transcript = transcript & "|stateLive=" & app.State("qty")
    transcript = transcript & "|noChangeYet=" & gChangeCount
    ReDimUI.OverridePointer 132, 36, False
    ReDimUI.PumpOnce
    transcript = transcript & "|releaseChange=" & gChangeCount
    ReDimUI.DispatchShape "rdm_wid46_qty__plus"
    transcript = transcript & "|releaseClickSwallowed=" & app.Stepper("qty").CurrentValue

    Sleep 450
    ReDimUI.OverridePointer 132, 36, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 132, 36, False
    ReDimUI.PumpOnce
    ReDimUI.DispatchShape "rdm_wid46_qty__plus"
    transcript = transcript & "|tapSteps=" & app.Stepper("qty").CurrentValue & ":" & gChangeCount

    ReDimUI.OverridePointer 335, 172, True
    ReDimUI.PumpOnce
    Sleep 1100
    ReDimUI.PumpOnce
    Sleep 450
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 335, 172, False
    ReDimUI.PumpOnce
    transcript = transcript & "|arrowPages=" & RowItem(host, "rdm_wid46_tl__al1")

    ' A held table pager pages on, and a held calendar arrow turns month
    ' after month.
    app.Table("tb").AtRect(24, 220, 300, 120).Columns "Row"
    For itemNo = 1 To 60
        app.Table("tb").AddRow "Row " & Format$(itemNo, "00")
    Next itemNo
    app.DatePicker("due").AtRect(360, 220, 150, 24).PickDate DateSerial(2026, 3, 10)
    HoldOnPart host, "rdm_wid46_tb__tn"
    transcript = transcript & "|tablePagesOn=" & CStr(TableRowShown(host, "rdm_wid46_tb__tr1") _
        <> "Row 01")
    Sleep 450
    ReDimUI.DispatchShape "rdm_wid46_due"
    HoldOnPart host, "rdm_wid46_due__cn"
    transcript = transcript & "|monthsTurn=" & CStr( _
        host.Shapes("rdm_wid46_due__ch").TextFrame2.TextRange.Text <> _
            Format$(DateSerial(2026, 4, 1), "mmmm yyyy") And _
        host.Shapes("rdm_wid46_due__ch").TextFrame2.TextRange.Text <> _
            Format$(DateSerial(2026, 3, 1), "mmmm yyyy"))
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestHoldRepeat = transcript
End Function

' Transfer gestures: a double click on a row moves it across; a press on
' a row dragged onto the other panel outlines that panel and the release
' there moves the row; a press dragged along a panel selects the rows it
' covers and swallows the release click.
Public Function TestTransferGestures() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid47")
    app.TransferList("tl").AtRect(24, 24, 300, 200) _
        .Items("A", "B", "C", "D", "E", "F", "G", "H").WritesTo("chosen") _
        .OnChange "TestReDimWidgets.RecordChange"
    app.Render

    ReDimUI.DispatchShape "rdm_wid47_tl__al2"
    ReDimUI.DispatchShape "rdm_wid47_tl__al2"
    transcript = "doubleClickMoves=" & app.State("chosen") & ":" & _
        RowItem(host, "rdm_wid47_tl__al2") & ":" & gChangeCount

    Sleep 600
    ReDimUI.OverridePointer 60, 55, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 240, 100, True
    ReDimUI.PumpOnce
    transcript = transcript & "|dropOutlined=" & _
        CStr(host.Shapes("rdm_wid47_tl__rp").Line.ForeColor.RGB = app.Theme.PrimaryColor)
    ReDimUI.OverridePointer 240, 100, False
    ReDimUI.PumpOnce
    transcript = transcript & "|dragMoves=" & app.State("chosen") & ":" & gChangeCount
    transcript = transcript & "|outlineCleared=" & _
        CStr(host.Shapes("rdm_wid47_tl__rp").Line.ForeColor.RGB = app.Theme.BorderColor)

    ReDimUI.OverridePointer 60, 55, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 60, 95, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 60, 95, False
    ReDimUI.PumpOnce
    ReDimUI.DispatchShape "rdm_wid47_tl__al3"
    transcript = transcript & "|rangeSelected=" & _
        CStr(RowChecked(host, "rdm_wid47_tl__al1") _
            And RowChecked(host, "rdm_wid47_tl__al2") _
            And RowChecked(host, "rdm_wid47_tl__al3") _
            And Not RowChecked(host, "rdm_wid47_tl__al4"))
    Sleep 450
    ReDimUI.DispatchShape "rdm_wid47_tl__mvr"
    transcript = transcript & "|rangeMoves=" & app.State("chosen")
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestTransferGestures = transcript
End Function

' Label/value items: a picked item writes its value to WritesTo from a
' select, a radio group, a combo pick, a check list, and a transfer list,
' while a combo's free text writes the text; ItemValueAt reads a value;
' a replaced list drops its values, and chosen transfer items keep theirs.
Public Function TestItemValues() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid48")
    app.SelectBox("sz").AtRect(24, 24, 140, 22) _
        .ItemsFrom(Array("Small", "Medium", "Large"), Array(10, 20, 30)).WritesTo "size"
    app.RadioGroup("rg").AtRect(24, 60, 140, 40).WritesTo "tier"
    app.RadioGroup("rg").AddItem "Free", , "F"
    app.RadioGroup("rg").AddItem "Pro", , "P"
    app.ComboBox("cb").AtRect(24, 120, 140, 22) _
        .ItemsFrom(Array("Red", "Green"), Array("#f00", "#0f0")).WritesTo "hue"
    app.CheckList("ck").AtRect(200, 24, 140, 80) _
        .ItemsFrom(Array("A", "B", "C"), Array(1, 2, 3)).WritesTo "picked"
    app.TransferList("tl").AtRect(200, 120, 300, 120) _
        .ItemsFrom(Array("X", "Y"), Array("x1", "y1")).WritesTo "moved"
    app.Render

    ReDimUI.DispatchShape "rdm_wid48_sz"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid48_sz__opt2"
    transcript = "selectValue=" & app.State("size") & ":" & TypeName(app.State("size"))
    transcript = transcript & "|itemValueAt=" & app.SelectBox("sz").ItemValueAt(3) & _
        ":" & app.SelectBox("sz").ItemValueAt(9)

    ReDimUI.DispatchShape "rdm_wid48_rg__t2"
    transcript = transcript & "|radioValue=" & app.State("tier")

    ReDimUI.DispatchShape "rdm_wid48_cb"
    ReDimUI.DispatchShape "rdm_wid48_cb__opt1"
    transcript = transcript & "|comboPick=" & app.State("hue")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid48_cb"
    BackspaceAll app.ComboBox("cb")
    TypeText "Blue"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|comboFreeText=" & app.State("hue")

    ReDimUI.DispatchShape "rdm_wid48_ck__t1"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid48_ck__t3"
    transcript = transcript & "|checkValues=" & app.State("picked")

    ReDimUI.DispatchShape "rdm_wid48_tl__al2"
    ReDimUI.DispatchShape "rdm_wid48_tl__mvr"
    transcript = transcript & "|transferValues=" & app.State("moved") & ":" & _
        app.TransferList("tl").ChosenValueAt(1)

    app.SelectBox("sz").Items "One", "Two"
    transcript = transcript & "|replacedDropsValues=" & app.SelectBox("sz").ItemValueAt(1)
    app.TransferList("tl").ItemsFrom Array("Z")
    transcript = transcript & "|chosenKeepValue=" & app.TransferList("tl").ChosenValueAt(1)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestItemValues = transcript
End Function

' SelectBox groups and disabled items: a header reads bold, muted, and
' flush left, a disabled item reads muted, clicks on either do nothing,
' the keys and type-ahead pass over both, a header is never the value,
' and both marks follow their items through an insert.
Public Function TestSelectGroups() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid49")
    app.SelectBox("dish").AtRect(24, 24, 160, 22).AddGroup "Fruit"
    app.SelectBox("dish").AddItem "Apple"
    app.SelectBox("dish").AddItem "Banana"
    app.SelectBox("dish").AddGroup "Vegetables"
    app.SelectBox("dish").AddItem "Carrot"
    app.SelectBox("dish").AddItem "Daikon"
    app.SelectBox("dish").ItemEnabled(3, False).WritesTo "dish"
    app.Render

    ReDimUI.DispatchShape "rdm_wid49_dish"
    With host.Shapes("rdm_wid49_dish__opt1").TextFrame2.TextRange
        transcript = "headerLook=" & CStr(.Text = "Fruit" _
            And .Characters(1, 1).Font.Bold = msoTrue _
            And .Font.Fill.ForeColor.RGB = app.Theme.OnMutedColor)
    End With
    transcript = transcript & "|disabledLook=" & _
        CStr(InkOf(host, "rdm_wid49_dish__opt3") = app.Theme.OnMutedColor _
            And RowItem(host, "rdm_wid49_dish__opt3") = "Banana" _
            And host.Shapes("rdm_wid49_dish__opt2").TextFrame2.TextRange _
                .Characters(2, 1).Font.Bold = msoFalse)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid49_dish__opt1"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid49_dish__opt3"
    transcript = transcript & "|clicksIgnored=" & _
        CStr(ShapeExists(host, "rdm_wid49_dish__opt1") And Not app.HasState("dish"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid49_dish__opt2"
    transcript = transcript & "|picked=" & app.State("dish")

    app.SelectBox("dish").Focus
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|downSkips=" & app.State("dish")
    RdxKeyChar "{UP}"
    transcript = transcript & "|upSkips=" & app.State("dish")
    RdxKeyChar "{HOME}"
    transcript = transcript & "|homeSkipsHeader=" & app.SelectBox("dish").CurrentValue
    RdxKeyChar "b"
    transcript = transcript & "|typeAheadSkips=" & app.State("dish")
    RdxKeyChar "{ESC}"

    app.SelectBox("dish").Value 1
    transcript = transcript & "|headerNotValue=" & app.SelectBox("dish").CurrentValue
    app.SelectBox("dish").AddItem "Avocado", 2
    transcript = transcript & "|marksShift=" & _
        CStr(app.SelectBox("dish").IsItemEnabled(3) _
            And Not app.SelectBox("dish").IsItemEnabled(4) _
            And Not app.SelectBox("dish").IsItemEnabled(5) _
            And app.SelectBox("dish").ItemTextAt(5) = "Vegetables")
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestSelectGroups = transcript
End Function

' Reorder: the chosen panel's arrows move its selected rows a place as a
' block, the selection going with them, and stop at the ends without a
' change; with nothing selected Alt+Up moves the cursor's row, and the
' cursor goes with it. Each move writes the new order and fires OnChange.
Public Function TestTransferReorder() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid50")
    app.TransferList("tl").AtRect(24, 24, 300, 160) _
        .ChosenFrom(Array("A", "B", "C", "D")).WritesTo("order") _
        .OnChange "TestReDimWidgets.RecordChange"
    app.TransferList("tl").Reorderable
    app.Render
    transcript = "arrowsDrawn=" & _
        CStr(ShapeExists(host, "rdm_wid50_tl__mvu") And ShapeExists(host, "rdm_wid50_tl__mvd"))

    ReDimUI.DispatchShape "rdm_wid50_tl__cl3"
    ReDimUI.DispatchShape "rdm_wid50_tl__cl4"
    ReDimUI.DispatchShape "rdm_wid50_tl__mvu"
    transcript = transcript & "|upOnce=" & app.State("order")
    ReDimUI.DispatchShape "rdm_wid50_tl__mvu"
    ReDimUI.DispatchShape "rdm_wid50_tl__mvu"
    transcript = transcript & "|stopsAtTop=" & app.State("order") & ":" & gChangeCount
    transcript = transcript & "|selectionFollows=" & _
        CStr(RowChecked(host, "rdm_wid50_tl__cl1") And RowChecked(host, "rdm_wid50_tl__cl2") _
            And Not RowChecked(host, "rdm_wid50_tl__cl3"))
    ReDimUI.DispatchShape "rdm_wid50_tl__mvd"
    transcript = transcript & "|downMoves=" & app.State("order")

    ReDimUI.DispatchShape "rdm_wid50_tl__cl2"
    ReDimUI.DispatchShape "rdm_wid50_tl__cl3"
    app.TransferList("tl").Focus
    RdxKeyChar "{RIGHT}"
    RdxKeyChar "{END}"
    RdxKeyChar "{ALTUP}"
    transcript = transcript & "|altUpMovesCursorRow=" & app.State("order")
    RdxKeyChar "{ALTUP}"
    transcript = transcript & "|cursorFollows=" & app.State("order")
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestTransferReorder = transcript
End Function

' Type-to-filter: typing while a transfer list has the keys filters the
' panel its cursor is in, the header shows the filter and the count, the
' keys walk the rows that show, Backspace widens, move-all moves only the
' rows that show, and Esc clears the filter before it leaves. A check
' list hides the rows the filter excludes, and select-all checks only
' the rows that show.
Public Function TestListFilter() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid51")
    app.TransferList("tl").AtRect(24, 24, 300, 160) _
        .Items("Apple", "Apricot", "Banana", "Blueberry", "Cherry").WritesTo "chosen"
    app.CheckList("ck").AtRect(360, 24, 160, 140) _
        .Items("Red", "Green", "Blue", "Gray").WritesTo "colors"
    app.Render

    app.TransferList("tl").Focus
    TypeText "ap"
    transcript = "headerShowsFilter=" & _
        host.Shapes("rdm_wid51_tl__hl").TextFrame2.TextRange.Text
    transcript = transcript & "|rowsFiltered=" & RowItem(host, "rdm_wid51_tl__al1") & "," & _
        RowItem(host, "rdm_wid51_tl__al2") & ":" & CStr(ShapeExists(host, "rdm_wid51_tl__al3"))
    RdxKeyChar "{DOWN}"
    RdxKeyChar " "
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keysOnShownRows=" & app.State("chosen")
    RdxKeyChar "{BS}"
    transcript = transcript & "|backspaceWidens=" & _
        host.Shapes("rdm_wid51_tl__hl").TextFrame2.TextRange.Text
    ReDimUI.DispatchShape "rdm_wid51_tl__mvar"
    transcript = transcript & "|moveAllShown=" & app.State("chosen")
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escClears=" & _
        host.Shapes("rdm_wid51_tl__hl").TextFrame2.TextRange.Text & ":" & _
        CStr(ReDimUI.HasKeyboardFocus)
    RdxKeyChar "{ESC}"

    app.CheckList("ck").Focus
    TypeText "gr"
    transcript = transcript & "|checkHeader=" & _
        host.Shapes("rdm_wid51_ck__mt").TextFrame2.TextRange.Text
    transcript = transcript & "|checkRowsHidden=" & _
        CStr(host.Shapes("rdm_wid51_ck").Visible = msoFalse _
            And host.Shapes("rdm_wid51_ck__b2").Visible = msoTrue _
            And host.Shapes("rdm_wid51_ck__b3").Visible = msoFalse _
            And host.Shapes("rdm_wid51_ck__b2").Top < host.Shapes("rdm_wid51_ck__b4").Top _
            And host.Shapes("rdm_wid51_ck__b4").Top < host.Shapes("rdm_wid51_ck__b3").Top)
    ReDimUI.DispatchShape "rdm_wid51_ck__mt"
    transcript = transcript & "|selectAllShown=" & app.State("colors")
    RdxKeyChar "{ESC}"
    transcript = transcript & "|allBack=" & _
        CStr(host.Shapes("rdm_wid51_ck").Visible = msoTrue _
            And host.Shapes("rdm_wid51_ck__mt").TextFrame2.TextRange.Text = "Select all (2/4)")
    RdxKeyChar "{ESC}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestListFilter = transcript
End Function

' Field adornments: a caption above with the required mark in the danger
' color, a hint below that gives way to the Required message an empty
' commit shows and comes back once the text passes, and an error set from
' outside with its danger border. A Skeleton draws its bars, the last one
' short, and pulses.
Public Function TestAdornments() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim firstFill As Long
    Dim transcript As String

    ReDimUI.AutoPump False
    ReDimUI.ReduceMotion False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid52")
    app.TextInput("email").AtRect(24, 40, 200, 22).Caption("Email").Hint("We never share it") _
        .Required
    app.TextInput("user").AtRect(24, 110, 200, 22).Caption "User name"
    app.SelectBox("size").AtRect(260, 40, 140, 22).Items("S", "M").Caption "Size"
    app.Skeleton("sk").AtRect(260, 100, 160, 60).SkeletonLines 3
    app.Render

    With host.Shapes("rdm_wid52_email__fc")
        transcript = "captionAbove=" & CStr(.TextFrame2.TextRange.Text = "Email *" _
            And .Top + .Height <= 40 _
            And .TextFrame2.TextRange.Characters(7, 1).Font.Fill.ForeColor.RGB _
                = app.Theme.DangerColor)
    End With
    transcript = transcript & "|hintBelow=" & _
        CStr(host.Shapes("rdm_wid52_email__fh").TextFrame2.TextRange.Text = "We never share it" _
            And host.Shapes("rdm_wid52_email__fh").Top >= 62)
    transcript = transcript & "|selectCaption=" & CStr(ShapeExists(host, "rdm_wid52_size__fc"))

    ReDimUI.DispatchShape "rdm_wid52_email"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|requiredShown=" & _
        CStr(host.Shapes("rdm_wid52_email__me").TextFrame2.TextRange.Text = "Required" _
            And Not ShapeExists(host, "rdm_wid52_email__fh") _
            And host.Shapes("rdm_wid52_email").Line.ForeColor.RGB = app.Theme.DangerColor)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid52_email"
    TypeText "a@b"
    transcript = transcript & "|requiredClears=" & _
        CStr(Not ShapeExists(host, "rdm_wid52_email__me") _
            And ShapeExists(host, "rdm_wid52_email__fh"))
    RdxKeyChar "{ENTER}"

    app.TextInput("user").ErrorText "Taken"
    transcript = transcript & "|errorTextShown=" & _
        CStr(host.Shapes("rdm_wid52_user__me").TextFrame2.TextRange.Text = "Taken" _
            And host.Shapes("rdm_wid52_user").Line.ForeColor.RGB = app.Theme.DangerColor)
    app.TextInput("user").ErrorText ""
    transcript = transcript & "|errorTextCleared=" & _
        CStr(Not ShapeExists(host, "rdm_wid52_user__me"))

    transcript = transcript & "|skeletonBars=" & _
        CStr(ShapeExists(host, "rdm_wid52_sk__sk3") _
            And host.Shapes("rdm_wid52_sk__sk3").Width < host.Shapes("rdm_wid52_sk__sk1").Width)
    firstFill = host.Shapes("rdm_wid52_sk__sk1").Fill.ForeColor.RGB
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    ReDimUI.PumpOnce
    transcript = transcript & "|skeletonPulses=" & _
        CStr(host.Shapes("rdm_wid52_sk__sk1").Fill.ForeColor.RGB <> firstFill)
    app.Skeleton("sk").Visible False
    RdxReleaseKeys
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestAdornments = transcript
End Function

' Tabs: the first tab shows its panel and hides the rest. A click on a
' tab switches panels, writes the tab's text, and fires OnChange once.
' Visible still decides a control on its own tab, and waits while the tab
' hides it. The keys walk the tabs, the pointer tints the tab under it, a
' focused field on a panel that hides commits, a hidden Tabs control
' hides every panel, and removing it shows them all.
Public Function TestTabs() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid53")
    app.Tabs("tabs").AtRect(24, 24, 360, 32).Items("General", "Advanced", "About") _
        .WritesTo "tabState"
    app.Tabs("tabs").OnChange "TestReDimWidgets.RecordChange"
    app.TextInput("name").AtRect(24, 70, 200, 22).WritesTo("nameState").OnTab "tabs", 1
    app.TickBox("beta").AtRect(24, 70, 200, 18).Text("Beta").OnTab "tabs", 2
    app.Label("later").AtRect(24, 100, 200, 20).Text("Soon").OnTab("tabs", 2).Visible False
    app.Label("ver").AtRect(24, 70, 200, 20).Text("Version 1").OnTab "tabs", 3
    app.Render

    transcript = "firstPanel=" & CStr(host.Shapes("rdm_wid53_name").Visible = msoTrue _
        And host.Shapes("rdm_wid53_beta").Visible = msoFalse _
        And host.Shapes("rdm_wid53_ver").Visible = msoFalse _
        And host.Shapes("rdm_wid53_later").Visible = msoFalse)
    transcript = transcript & "|firstBold=" & CStr( _
        host.Shapes("rdm_wid53_tabs__tb1").TextFrame2.TextRange.Font.Bold = msoTrue _
        And host.Shapes("rdm_wid53_tabs__tb2").TextFrame2.TextRange.Font.Bold = msoFalse)
    With host.Shapes("rdm_wid53_tabs__ti")
        transcript = transcript & "|barUnderFirst=" & CStr( _
            .Left >= host.Shapes("rdm_wid53_tabs__tb1").Left _
            And .Left + .Width <= host.Shapes("rdm_wid53_tabs__tb2").Left _
            And .Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    End With

    ReDimUI.DispatchShape "rdm_wid53_tabs__tb2"
    transcript = transcript & "|clickSwitches=" & CStr( _
        host.Shapes("rdm_wid53_name").Visible = msoFalse _
        And host.Shapes("rdm_wid53_beta").Visible = msoTrue _
        And app.Tabs("tabs").CurrentValue = 2)
    transcript = transcript & "|writes=" & CStr(app.State("tabState"))
    transcript = transcript & "|fires=" & gChangeCount
    transcript = transcript & "|hiddenStays=" & _
        CStr(host.Shapes("rdm_wid53_later").Visible = msoFalse)
    transcript = transcript & "|altText=" & CStr(InStr(1, _
        host.Shapes("rdm_wid53_tabs").AlternativeText, "Advanced selected, tab 2 of 3") > 0)

    app.Label("later").Visible True
    app.Label("ver").Visible False
    transcript = transcript & "|visibleOnTab=" & _
        CStr(host.Shapes("rdm_wid53_later").Visible = msoTrue)

    app.Tabs("tabs").Focus
    RdxKeyChar "{RIGHT}"
    transcript = transcript & "|keyRight=" & CStr(app.Tabs("tabs").CurrentValue = 3 _
        And host.Shapes("rdm_wid53_beta").Visible = msoFalse _
        And host.Shapes("rdm_wid53_ver").Visible = msoFalse)
    app.Label("ver").Visible True
    transcript = transcript & "|visibleWaited=" & _
        CStr(host.Shapes("rdm_wid53_ver").Visible = msoTrue)
    RdxKeyChar "{RIGHT}"
    transcript = transcript & "|keyWraps=" & CStr(app.Tabs("tabs").CurrentValue)
    RdxKeyChar "{END}"
    transcript = transcript & "|keyEnd=" & CStr(app.Tabs("tabs").CurrentValue)
    RdxKeyChar "{HOME}"
    RdxReleaseKeys

    app.PointerEffects
    ReDimUI.OverridePointer host.Shapes("rdm_wid53_tabs__tb3").Left + 10, 38
    ReDimUI.PumpOnce
    transcript = transcript & "|hoverTints=" & CStr( _
        host.Shapes("rdm_wid53_tabs__tb3").Fill.Visible = msoTrue _
        And host.Shapes("rdm_wid53_tabs__tb2").Fill.Visible = msoFalse)
    ReDimUI.ClearPointerOverride
    app.PointerEffects False

    ReDimUI.DispatchShape "rdm_wid53_name"
    TypeText "Ada"
    app.Tabs("tabs").Value 2
    transcript = transcript & "|hidingCommits=" & CStr(app.State("nameState") = "Ada" _
        And Not ReDimUI.IsComponentFocused("wid53", "name") _
        And host.Shapes("rdm_wid53_name").Visible = msoFalse)

    app.Tabs("tabs").Visible False
    transcript = transcript & "|tabsHidden=" & CStr( _
        host.Shapes("rdm_wid53_beta").Visible = msoFalse _
        And host.Shapes("rdm_wid53_tabs__tb1").Visible = msoFalse)
    app.Tabs("tabs").Visible True
    transcript = transcript & "|tabsBack=" & CStr( _
        host.Shapes("rdm_wid53_beta").Visible = msoTrue _
        And host.Shapes("rdm_wid53_name").Visible = msoFalse)

    app.Tabs("tabs").Remove
    transcript = transcript & "|removedShowsAll=" & CStr( _
        host.Shapes("rdm_wid53_name").Visible = msoTrue _
        And host.Shapes("rdm_wid53_beta").Visible = msoTrue _
        And host.Shapes("rdm_wid53_ver").Visible = msoTrue)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestTabs = transcript
End Function

' DatePicker: the face shows its placeholder muted, then the date in the
' format asked for; PickDate writes nothing and fires nothing. A click
' opens a calendar on the date's month with the date filled; a day
' outside DateRange does not pick, the arrow turns the month, and a day
' writes a Date, fires OnChange once, and closes. The keys open, walk,
' and pick; Esc closes. A press on the calendar keeps it, a press off it
' closes it, and the pointer tints the day under it.
Public Function TestDatePicker() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim firstCell As Long
    Dim pickedBefore As Date
    Dim gridStart As Date

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid54")
    app.DatePicker("due").AtRect(24, 24, 150, 24).Text("Pick a date").WritesTo "dueState"
    app.DatePicker("due").OnChange "TestReDimWidgets.RecordChange"
    app.DatePicker("due").DateRange DateSerial(2026, 9, 5), DateSerial(2026, 12, 31)
    app.Render

    With host.Shapes("rdm_wid54_due").TextFrame2.TextRange
        transcript = "placeholder=" & CStr(.Text = "Pick a date" _
            And .Font.Fill.ForeColor.RGB = app.Theme.OnMutedColor)
    End With
    app.DatePicker("due").DateFormat("yyyy-mm-dd").PickDate DateSerial(2026, 9, 22)
    transcript = transcript & "|faceFormat=" & _
        host.Shapes("rdm_wid54_due").TextFrame2.TextRange.Text
    transcript = transcript & "|picked=" & Format$(app.DatePicker("due").PickedDate, "yyyy-mm-dd")
    transcript = transcript & "|programSilent=" & _
        CStr(gChangeCount = 0 And Format$(app.StateOrDefault("dueState", 0), "yyyy-mm-dd") = "2026-09-22")

    ReDimUI.DispatchShape "rdm_wid54_due"
    transcript = transcript & "|opens=" & CStr(ShapeExists(host, "rdm_wid54_due__cr6") _
        And host.Shapes("rdm_wid54_due__ch").TextFrame2.TextRange.Text _
            = Format$(DateSerial(2026, 9, 1), "mmmm yyyy"))
    ' The 22nd reads in the accent's ink over the accent fill on its cell.
    firstCell = Weekday(DateSerial(2026, 9, 1), vbUseSystemDayOfWeek)
    transcript = transcript & "|dateFilled=" & CStr( _
        CalendarDayWords(host, "rdm_wid54_due", firstCell + 21) = "22" _
        And CalendarDayInk(host, "rdm_wid54_due", firstCell + 21) = app.Theme.OnPrimaryColor _
        And CalendarMarkOn(host, "rdm_wid54_due", "cs", firstCell + 21) _
        And host.Shapes("rdm_wid54_due__cs").Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    ReDimUI.DispatchShape "rdm_wid54_due__cd" & (firstCell + 2)
    transcript = transcript & "|outOfRangeKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid54_due__cr1") And gChangeCount = 0)
    ' A week or a mark deleted by hand comes back with the next repaint.
    host.Shapes("rdm_wid54_due__cr3").Delete
    host.Shapes("rdm_wid54_due__cs").Delete
    ReDimUI.DispatchShape "rdm_wid54_due__cn"
    transcript = transcript & "|nextMonth=" & CStr( _
        host.Shapes("rdm_wid54_due__ch").TextFrame2.TextRange.Text _
            = Format$(DateSerial(2026, 10, 1), "mmmm yyyy"))
    transcript = transcript & "|partsReturn=" & CStr(ShapeExists(host, "rdm_wid54_due__cs") _
        And UBound(Split(host.Shapes("rdm_wid54_due__cr3").TextFrame2.TextRange.Text, _
            vbTab)) = 7)
    transcript = transcript & "|fillBehindWeeks=" & CStr( _
        host.Shapes("rdm_wid54_due__cs").ZOrderPosition _
            < host.Shapes("rdm_wid54_due__cr1").ZOrderPosition)
    firstCell = Weekday(DateSerial(2026, 10, 1), vbUseSystemDayOfWeek)
    ReDimUI.DispatchShape "rdm_wid54_due__cd" & (firstCell + 14)
    transcript = transcript & "|dayPicks=" & CStr(Not ShapeExists(host, "rdm_wid54_due__cr1") _
        And VarType(app.State("dueState")) = vbDate _
        And gChangeCount = 1)
    transcript = transcript & "|written=" & Format$(app.State("dueState"), "yyyy-mm-dd")
    transcript = transcript & "|faceShows=" & _
        host.Shapes("rdm_wid54_due").TextFrame2.TextRange.Text

    app.DatePicker("due").Focus
    RdxKeyChar "{ALTDOWN}"
    transcript = transcript & "|keyOpens=" & CStr(ShapeExists(host, "rdm_wid54_due__cr1"))
    RdxKeyChar "{RIGHT}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{PGDN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keysPick=" & Format$(app.State("dueState"), "yyyy-mm-dd") & _
        "/" & gChangeCount
    RdxKeyChar "{ALTDOWN}"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escCloses=" & CStr(Not ShapeExists(host, "rdm_wid54_due__cr1"))
    ' The keys stop at DateRange's ends, so Enter always has a day to pick.
    RdxKeyChar "{ALTDOWN}"
    RdxKeyChar "{PGDN}"
    RdxKeyChar "{PGDN}"
    RdxKeyChar "{PGDN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keysStopAtLatest=" & Format$(app.State("dueState"), "yyyy-mm-dd")
    RdxKeyChar "{ALTDOWN}"
    RdxKeyChar "{PGUP}"
    RdxKeyChar "{PGUP}"
    RdxKeyChar "{PGUP}"
    RdxKeyChar "{PGUP}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keysStopAtEarliest=" & Format$(app.State("dueState"), "yyyy-mm-dd")
    ' The month arrows stop at DateRange too, and read muted there.
    RdxKeyChar "{ALTDOWN}"
    transcript = transcript & "|arrowsAtStart=" & CStr( _
        InkOf(host, "rdm_wid54_due__cp") = app.Theme.OnMutedColor And _
        InkOf(host, "rdm_wid54_due__cn") = app.Theme.OnSurfaceColor)
    ReDimUI.DispatchShape "rdm_wid54_due__cp"
    transcript = transcript & "|prevRefused=" & CStr( _
        host.Shapes("rdm_wid54_due__ch").TextFrame2.TextRange.Text _
            = Format$(DateSerial(2026, 9, 1), "mmmm yyyy"))
    ReDimUI.DispatchShape "rdm_wid54_due__cn"
    ReDimUI.DispatchShape "rdm_wid54_due__cn"
    ReDimUI.DispatchShape "rdm_wid54_due__cn"
    ReDimUI.DispatchShape "rdm_wid54_due__cn"
    transcript = transcript & "|nextStopsAtEnd=" & CStr( _
        host.Shapes("rdm_wid54_due__ch").TextFrame2.TextRange.Text _
            = Format$(DateSerial(2026, 12, 1), "mmmm yyyy") And _
        InkOf(host, "rdm_wid54_due__cn") = app.Theme.OnMutedColor)
    RdxKeyChar "{ESC}"
    RdxReleaseKeys

    Sleep 200
    ReDimUI.DispatchShape "rdm_wid54_due"
    ReDimUI.OverridePointer 124, 150
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|pressOnKeeps=" & CStr(ShapeExists(host, "rdm_wid54_due__cr1"))
    app.PointerEffects
    ReDimUI.OverridePointer host.Shapes("rdm_wid54_due__cr2").Left + 2 * 26 + 13, _
        host.Shapes("rdm_wid54_due__cr2").Top + 10
    ReDimUI.PumpOnce
    transcript = transcript & "|hoverTints=" & CStr( _
        CalendarMarkOn(host, "rdm_wid54_due", "cv", 10) _
        And host.Shapes("rdm_wid54_due__cv").Fill.Visible = msoTrue)
    app.PointerEffects False
    ReDimUI.OverridePointer 420, 300
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|pressOffCloses=" & CStr(Not ShapeExists(host, "rdm_wid54_due__cr1"))

    ' A click on a week line picks the day under the pointer. A week is one
    ' shape, so this is the path a real click on a day takes.
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid54_due"
    pickedBefore = app.DatePicker("due").PickedDate
    gridStart = DateSerial(Year(pickedBefore), Month(pickedBefore), 1)
    gridStart = gridStart - (Weekday(gridStart, vbUseSystemDayOfWeek) - 1)
    ReDimUI.OverridePointer host.Shapes("rdm_wid54_due__cr2").Left + 2 * 26 + 13, _
        host.Shapes("rdm_wid54_due__cr2").Top + 10
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid54_due__cr2"
    transcript = transcript & "|weekClickPicks=" & CStr( _
        app.DatePicker("due").PickedDate = gridStart + 9 _
        And Not ShapeExists(host, "rdm_wid54_due__cr1"))
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestDatePicker = transcript
End Function

' A calendar grid cell (1 to 42) as its week's line shows it: the digits,
' and their ink.
Private Function CalendarDayWords(ByVal host As Worksheet, ByVal pickerName As String, _
    ByVal cellNo As Long) As String
    Dim dayFields() As String

    dayFields = Split(host.Shapes(pickerName & "__cr" & ((cellNo - 1) \ 7 + 1)) _
        .TextFrame2.TextRange.Text, vbTab)
    CalendarDayWords = dayFields((cellNo - 1) Mod 7 + 1)
End Function

Private Function CalendarDayInk(ByVal host As Worksheet, ByVal pickerName As String, _
    ByVal cellNo As Long) As Long
    Dim dayFields() As String
    Dim fieldNo As Long
    Dim startAt As Long

    With host.Shapes(pickerName & "__cr" & ((cellNo - 1) \ 7 + 1)).TextFrame2.TextRange
        dayFields = Split(.Text, vbTab)
        startAt = 1
        For fieldNo = 0 To (cellNo - 1) Mod 7
            startAt = startAt + Len(dayFields(fieldNo)) + 1
        Next fieldNo
        CalendarDayInk = .Characters(startAt, Len(dayFields((cellNo - 1) Mod 7 + 1))) _
            .Font.Fill.ForeColor.RGB
    End With
End Function

' True when a day mark shows on a grid cell: visible, over the cell's
' column in the cell's week.
Private Function CalendarMarkOn(ByVal host As Worksheet, ByVal pickerName As String, _
    ByVal markName As String, ByVal cellNo As Long) As Boolean
    Dim weekShape As Shape

    Set weekShape = host.Shapes(pickerName & "__cr" & ((cellNo - 1) \ 7 + 1))
    With host.Shapes(pickerName & "__" & markName)
        CalendarMarkOn = (.Visible = msoTrue) _
            And (Abs(.Left - (weekShape.Left + ((cellNo - 1) Mod 7) * 26 + 1)) < 0.5) _
            And (Abs(.Top - weekShape.Top) < 0.5)
    End With
End Function

' Table: headers over the columns, a number column aligned right, rows in
' the order they came, formats applied. A header click sorts ascending
' and a second click descending, keeping equal rows in order. A row click
' selects, writes the first cell, and fires OnChange once. The footer
' pages when the rows outgrow the table, and the keys move the selection
' and scroll it into view. TableFrom reads a range under its header row,
' and SortBy puts empty cells last.
Public Function TestTable() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim rowNo As Long

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid55")
    app.Table("orders").AtRect(24, 24, 360, 160).Columns("Item", "Qty", "Price") _
        .WritesTo "orderState"
    app.Table("orders").OnChange "TestReDimWidgets.RecordChange"
    app.Table("orders").AddRow "Pear", 3, 1.25
    app.Table("orders").AddRow "Apple", 12, 0.5
    app.Table("orders").AddRow "Fig", 3, 2
    app.Table("orders").AddRow "Kiwi", 7, 0.75
    app.Table("orders").ColumnFormat 3, "0.00"
    app.Render

    transcript = "heads=" & host.Shapes("rdm_wid55_orders__th1").TextFrame2.TextRange.Text & _
        "/" & host.Shapes("rdm_wid55_orders__th2").TextFrame2.TextRange.Text
    transcript = transcript & "|numberRight=" & CStr( _
        host.Shapes("rdm_wid55_orders__th2").TextFrame2.TextRange.ParagraphFormat.Alignment _
            = msoAlignRight _
        And host.Shapes("rdm_wid55_orders__th1").TextFrame2.TextRange.ParagraphFormat.Alignment _
            = msoAlignLeft)
    transcript = transcript & "|firstRow=" & TableRowShown(host, "rdm_wid55_orders__tr1")
    transcript = transcript & "|noFooter=" & CStr(Not ShapeExists(host, "rdm_wid55_orders__tf"))

    ReDimUI.DispatchShape "rdm_wid55_orders__th2"
    transcript = transcript & "|sortUp=" & TableRowShown(host, "rdm_wid55_orders__tr1") & _
        "," & TableRowShown(host, "rdm_wid55_orders__tr2") & _
        "," & TableRowShown(host, "rdm_wid55_orders__tr4")
    transcript = transcript & "|arrowUp=" & CStr( _
        host.Shapes("rdm_wid55_orders__th2").TextFrame2.TextRange.Text = "Qty " & ChrW(9650))
    ReDimUI.DispatchShape "rdm_wid55_orders__th2"
    transcript = transcript & "|sortDown=" & TableRowShown(host, "rdm_wid55_orders__tr1") & _
        "," & TableRowShown(host, "rdm_wid55_orders__tr3") & _
        "," & TableRowShown(host, "rdm_wid55_orders__tr4")
    transcript = transcript & "|sortSilent=" & CStr(gChangeCount = 0)

    ReDimUI.DispatchShape "rdm_wid55_orders__tr2"
    transcript = transcript & "|rowPicks=" & CStr(app.State("orderState")) & "/" & _
        CStr(app.Table("orders").CurrentValue) & "/" & gChangeCount
    transcript = transcript & "|rowFilled=" & CStr( _
        host.Shapes("rdm_wid55_orders__tr2").Fill.Visible = msoTrue _
        And host.Shapes("rdm_wid55_orders__tr2").Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    ReDimUI.DispatchShape "rdm_wid55_orders__tr2"
    transcript = transcript & "|repickSilent=" & CStr(gChangeCount = 1)

    For rowNo = 1 To 10
        app.Table("orders").AddRow "Item " & rowNo, rowNo, rowNo / 4
    Next rowNo
    transcript = transcript & "|footer=" & _
        host.Shapes("rdm_wid55_orders__tf").TextFrame2.TextRange.Text
    ReDimUI.DispatchShape "rdm_wid55_orders__tn"
    transcript = transcript & "|paged=" & _
        host.Shapes("rdm_wid55_orders__tf").TextFrame2.TextRange.Text

    app.Table("orders").Focus
    RdxKeyChar "{HOME}"
    transcript = transcript & "|keyHome=" & _
        host.Shapes("rdm_wid55_orders__tf").TextFrame2.TextRange.Text & "/" & _
        CStr(app.State("orderState"))
    RdxKeyChar "{END}"
    transcript = transcript & "|keyEnd=" & _
        host.Shapes("rdm_wid55_orders__tf").TextFrame2.TextRange.Text & "/" & _
        CStr(app.State("orderState"))
    RdxKeyChar "{UP}"
    transcript = transcript & "|keyUp=" & CStr(app.State("orderState")) & "/" & gChangeCount
    RdxReleaseKeys

    host.Range("H1:I4").Value = Array("Name", "Score")
    host.Range("H2").Value = "Ann"
    host.Range("I2").Value = 90
    host.Range("H3").Value = "Bob"
    host.Range("I3").Value = 85
    host.Range("H4").Value = "Cy"
    host.Range("I4").ClearContents
    app.Table("scores").AtRect(24, 220, 240, 120).TableFrom host.Range("H1:I4")
    app.Table("scores").SortBy 2
    transcript = transcript & "|fromRange=" & _
        host.Shapes("rdm_wid55_scores__th1").TextFrame2.TextRange.Text & "/" & _
        app.Table("scores").RowCount & "/" & app.Table("scores").CellValue(2, 2)
    transcript = transcript & "|blanksLast=" & TableRowShown(host, "rdm_wid55_scores__tr1") & _
        "," & TableRowShown(host, "rdm_wid55_scores__tr3")
    ReDimUI.AutoPump True
    TestTable = transcript
End Function

' A table row's text with its tabs as spaces and outer spaces trimmed.
' Holds the left button on a part's middle past the slowest keyboard
' repeat delay and one repeat interval, then lets go. No click follows,
' so whatever changed came from the hold's repeats.
Private Sub HoldOnPart(ByVal host As Worksheet, ByVal shapeName As String)
    Dim pointsX As Double
    Dim pointsY As Double

    With host.Shapes(shapeName)
        pointsX = .Left + .Width / 2
        pointsY = .Top + .Height / 2
    End With
    ReDimUI.OverridePointer pointsX, pointsY, True
    ReDimUI.PumpOnce
    Sleep 1100
    ReDimUI.PumpOnce
    Sleep 450
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer pointsX, pointsY, False
    ReDimUI.PumpOnce
End Sub

Private Function TableRowShown(ByVal host As Worksheet, ByVal shapeName As String) As String
    TableRowShown = Trim$(Replace(host.Shapes(shapeName).TextFrame2.TextRange.Text, vbTab, " "))
End Function

' Icons: a button's icon leads its text in the icon font while the text
' keeps the theme font; an icon alone names the button for screen
' readers; a label takes one too. IconGlyph maps names, passes one
' character through, and refuses an unknown name; Icon "" takes the icon
' off and gives the text its font back.
Public Function TestIcons() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim refused As Boolean

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid56")
    app.Button("save").AtRect(24, 24, 120, 30).Icon("Save").Text "Save"
    app.Button("trash").AtRect(160, 24, 40, 30).Icon "Delete"
    app.Label("note").AtRect(24, 70, 200, 20).Icon("Info").Text "Saved drafts"
    app.Render

    With host.Shapes("rdm_wid56_save").TextFrame2.TextRange
        transcript = "iconText=" & CStr(.Text = ReDimUI.IconGlyph("Save") & "  Save" _
            And .Characters(1, 1).Font.Name = ReDimUI.IconFont _
            And .Characters(4, 4).Font.Name = app.Theme.FontName)
    End With
    transcript = transcript & "|iconAlone=" & CStr( _
        host.Shapes("rdm_wid56_trash").TextFrame2.TextRange.Text = ReDimUI.IconGlyph("Delete") _
        And host.Shapes("rdm_wid56_trash").AlternativeText = "Delete, button")
    transcript = transcript & "|labelIcon=" & CStr( _
        host.Shapes("rdm_wid56_note").TextFrame2.TextRange.Characters(1, 1).Font.Name _
            = ReDimUI.IconFont)
    transcript = transcript & "|glyph=" & CStr(AscW(ReDimUI.IconGlyph("Add")) And &HFFFF&) & _
        "/" & CStr(ReDimUI.IconGlyph("x") = "x")
    On Error Resume Next
    ReDimUI.IconGlyph "NoSuchIcon"
    refused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    transcript = transcript & "|unknownRefused=" & CStr(refused)
    app.Button("save").Icon ""
    With host.Shapes("rdm_wid56_save").TextFrame2.TextRange
        transcript = transcript & "|iconOff=" & CStr(.Text = "Save" _
            And .Characters(1, 1).Font.Name = app.Theme.FontName)
    End With
    transcript = transcript & "|iconFont=" & ReDimUI.IconFont
    ReDimUI.AutoPump True
    TestIcons = transcript
End Function

' The Windows look: ThemeSystem takes the dark theme under dark mode with
' the accent made readable, and the light one with its own primary when
' there is no accent. An app that follows the look re-themes when it
' changes, and one that stops keeps its theme. Reading the real registry
' works too.
Public Function TestSystemTheme() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim themeValue As ReDimUI
    Dim transcript As String

    ReDimUI.AutoPump False
    ReDimUI.OverrideSystemLook True, RGB(0, 120, 212)
    Set themeValue = ReDimUI.ThemeSystem
    transcript = "darkSurface=" & CStr(themeValue.SurfaceColor = ReDimUI.ThemeDark.SurfaceColor)
    transcript = transcript & "|accentReads=" & CStr( _
        ReDimUI.ContrastRatio(themeValue.PrimaryColor, themeValue.SurfaceColor) >= 3 _
        And ReDimUI.ContrastRatio(themeValue.OnPrimaryColor, themeValue.PrimaryColor) >= 4.5 _
        And themeValue.PrimaryColor <> ReDimUI.ThemeDark.PrimaryColor)
    ReDimUI.OverrideSystemLook False, -1
    Set themeValue = ReDimUI.ThemeSystem
    transcript = transcript & "|lightNoAccent=" & CStr( _
        themeValue.PrimaryColor = ReDimUI.ThemeLight.PrimaryColor _
        And themeValue.SurfaceColor = ReDimUI.ThemeLight.SurfaceColor)

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid57")
    app.Button("go").AtRect(24, 24, 100, 30).Text "Go"
    app.FollowSystemTheme
    app.Render
    transcript = transcript & "|followsLight=" & _
        CStr(app.Theme.SurfaceColor = ReDimUI.ThemeLight.SurfaceColor)
    ReDimUI.OverrideSystemLook True, RGB(0, 120, 212)
    ReDimUI.CheckSystemLook True
    transcript = transcript & "|switchesDark=" & CStr( _
        app.Theme.SurfaceColor = ReDimUI.ThemeDark.SurfaceColor _
        And host.Shapes("rdm_wid57_go").Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    app.FollowSystemTheme False
    ReDimUI.OverrideSystemLook False, -1
    ReDimUI.CheckSystemLook True
    transcript = transcript & "|stopsFollowing=" & _
        CStr(app.Theme.SurfaceColor = ReDimUI.ThemeDark.SurfaceColor)
    ReDimUI.ClearSystemLookOverride
    Set themeValue = ReDimUI.ThemeSystem
    transcript = transcript & "|realRead=" & CStr(Not themeValue Is Nothing)
    ReDimUI.AutoPump True
    TestSystemTheme = transcript
End Function

' Badges: a Badge pill widens to its text and takes its variant's fill.
' BadgeText puts a danger pill on a control's top-right corner, hides with
' the control, and "" takes it off.
Public Function TestBadges() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim shortWidth As Double

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid58")
    app.Badge("count").AtRect(24, 24, 0, 18).Text "3"
    app.Badge("status").AtRect(24, 60, 0, 18).Text("Overdue").Danger
    app.Button("inbox").AtRect(120, 24, 100, 30).Text("Inbox").BadgeText "12"
    app.Render
    shortWidth = host.Shapes("rdm_wid58_count").Width
    transcript = "widens=" & CStr(host.Shapes("rdm_wid58_status").Width > shortWidth _
        And shortWidth >= 18)
    transcript = transcript & "|tones=" & CStr( _
        host.Shapes("rdm_wid58_count").Fill.ForeColor.RGB = app.Theme.PrimaryColor _
        And host.Shapes("rdm_wid58_status").Fill.ForeColor.RGB = app.Theme.DangerColor)
    With host.Shapes("rdm_wid58_inbox__bd")
        transcript = transcript & "|cornerBadge=" & CStr(.TextFrame2.TextRange.Text = "12" _
            And .Left > host.Shapes("rdm_wid58_inbox").Left + 60 _
            And .Top < host.Shapes("rdm_wid58_inbox").Top _
            And .Fill.ForeColor.RGB = app.Theme.DangerColor)
    End With
    app.Button("inbox").Visible False
    transcript = transcript & "|hidesWith=" & CStr(Not ShapeExists(host, "rdm_wid58_inbox__bd"))
    app.Button("inbox").Visible True
    app.Button("inbox").BadgeText ""
    transcript = transcript & "|removed=" & CStr(Not ShapeExists(host, "rdm_wid58_inbox__bd"))
    ReDimUI.AutoPump True
    TestBadges = transcript
End Function

' Stack: members flow down in join order a gap apart, the caption row
' above a field and its note row below counted, and the stack grows to
' its content. A hidden member gives its place to the next and a shown
' one takes it back; a member that grows moves the ones after it. A stack
' across lines its members up left to right; Stretch widens the controls
' but not a stack inside. A hidden stack hides its members, and removing
' a member closes its gap.
Public Function TestStack() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid59")
    app.Stack("form").AtRect(24, 24, 260, 0).Gap(10).Stretch
    app.TextInput("name").Sized(100, 22).Caption("Name").InStack "form"
    app.TextInput("email").Sized(100, 22).Caption("Email").Hint("We never share it") _
        .InStack "form"
    app.Label("news").Sized(160, 18).Text("Newsletter").InStack "form"
    app.Stack("row").Across.Gap(8).InStack "form"
    app.Button("ok").Sized(80, 28).Text("OK").InStack "row"
    app.Button("cancel").Sized(80, 28).Text("Cancel").InStack "row"
    app.Render

    transcript = "tops=" & host.Shapes("rdm_wid59_name").Top & "," & _
        host.Shapes("rdm_wid59_email").Top & "," & host.Shapes("rdm_wid59_news").Top & _
        "," & host.Shapes("rdm_wid59_ok").Top
    transcript = transcript & "|stretched=" & host.Shapes("rdm_wid59_name").Width & "," & _
        host.Shapes("rdm_wid59_row").Width
    transcript = transcript & "|across=" & host.Shapes("rdm_wid59_ok").Left & "," & _
        host.Shapes("rdm_wid59_cancel").Left
    transcript = transcript & "|height=" & host.Shapes("rdm_wid59_form").Height

    app.TextInput("email").Visible False
    transcript = transcript & "|collapsed=" & host.Shapes("rdm_wid59_news").Top & "," & _
        host.Shapes("rdm_wid59_ok").Top & "," & host.Shapes("rdm_wid59_form").Height
    app.TextInput("email").Visible True
    transcript = transcript & "|restored=" & host.Shapes("rdm_wid59_news").Top & "," & _
        host.Shapes("rdm_wid59_ok").Top

    app.TextInput("name").Sized 100, 40
    transcript = transcript & "|grew=" & host.Shapes("rdm_wid59_email").Top & "," & _
        host.Shapes("rdm_wid59_news").Top

    app.Stack("row").Visible False
    transcript = transcript & "|stackHidden=" & CStr( _
        host.Shapes("rdm_wid59_ok").Visible = msoFalse _
        And host.Shapes("rdm_wid59_cancel").Visible = msoFalse) & "," & _
        host.Shapes("rdm_wid59_form").Height
    app.Stack("row").Visible True
    transcript = transcript & "|stackShown=" & CStr(host.Shapes("rdm_wid59_ok").Visible = msoTrue)

    app.Label("news").Remove
    transcript = transcript & "|removed=" & host.Shapes("rdm_wid59_ok").Top
    ReDimUI.AutoPump True
    TestStack = transcript
End Function

' Expander: closed by default with its panel hidden under a right-pointing
' chevron. A click opens it, shows the panel, writes the state, and fires
' OnChange, and in a stack the controls after it move down to make room;
' Left closes it and Right opens it from the keys.
Public Function TestExpander() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid60")
    app.Stack("page").AtRect(24, 24, 260, 0).Gap 8
    app.Expander("adv").Sized(260, 28).Text("Advanced").WritesTo("advOpen").InStack "page"
    app.Expander("adv").OnChange "TestReDimWidgets.RecordChange"
    app.TextInput("proxy").Sized(200, 22).InExpander("adv").InStack "page"
    app.Label("after").Sized(200, 18).Text("After").InStack "page"
    app.Render

    transcript = "closed=" & CStr(host.Shapes("rdm_wid60_proxy").Visible = msoFalse _
        And Left$(host.Shapes("rdm_wid60_adv").TextFrame2.TextRange.Text, 1) _
            = ReDimUI.IconGlyph("ChevronRight")) & "," & host.Shapes("rdm_wid60_after").Top
    ReDimUI.DispatchShape "rdm_wid60_adv"
    transcript = transcript & "|opens=" & CStr( _
        host.Shapes("rdm_wid60_proxy").Visible = msoTrue _
        And app.State("advOpen") = True And gChangeCount = 1 _
        And app.Expander("adv").IsExpanded _
        And Left$(host.Shapes("rdm_wid60_adv").TextFrame2.TextRange.Text, 1) _
            = ReDimUI.IconGlyph("ChevronDown")) & "," & host.Shapes("rdm_wid60_after").Top
    transcript = transcript & "|alt=" & CStr(InStr(1, _
        host.Shapes("rdm_wid60_adv").AlternativeText, "Advanced, expander, expanded") > 0)
    app.Expander("adv").Focus
    RdxKeyChar "{LEFT}"
    transcript = transcript & "|keyCloses=" & CStr( _
        host.Shapes("rdm_wid60_proxy").Visible = msoFalse And gChangeCount = 2) & "," & _
        host.Shapes("rdm_wid60_after").Top
    RdxKeyChar "{RIGHT}"
    transcript = transcript & "|keyOpens=" & CStr(host.Shapes("rdm_wid60_proxy").Visible = msoTrue)
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestExpander = transcript
End Function

' MenuButton: a button face in its variant's colors under the menu's own
' text. A click drops the commands with their icons in the gutter; a pick
' runs the command's handler with the menu as the sender and keeps the
' face's text; a disabled command runs nothing. The keys open the menu,
' walk it, and run the highlighted command.
Public Function TestMenuButton() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim farX As Double
    Dim farY As Double
    Dim idx As Long
    Dim narrowWidth As Double
    Dim longRow As String

    gCommandCount = 0
    gLastCommand = vbNullString
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid61")
    app.MenuButton("actions").AtRect(24, 24, 120, 28).Text("Actions") _
        .AddCommand("Export", "TestReDimWidgets.RecordCommand", "Download") _
        .AddCommand "Duplicate", "TestReDimWidgets.RecordCommand", "Copy"
    app.MenuButton("actions").AddGroup "Danger zone"
    app.MenuButton("actions").AddCommand("Delete", "TestReDimWidgets.RecordCommand", "Delete") _
        .ItemEnabled 4, False
    app.Render

    With host.Shapes("rdm_wid61_actions")
        transcript = "face=" & CStr(.TextFrame2.TextRange.Text = "Actions" _
            And .TextFrame2.TextRange.Font.Bold = msoTrue _
            And .Fill.ForeColor.RGB = app.Theme.SurfaceColor)
    End With
    ReDimUI.DispatchShape "rdm_wid61_actions"
    With host.Shapes("rdm_wid61_actions__opt1").TextFrame2.TextRange
        transcript = transcript & "|iconRow=" & CStr( _
            .Text = ReDimUI.IconGlyph("Download") & vbTab & "Export" _
            And .Characters(1, 1).Font.Name = ReDimUI.IconFont _
            And .Characters(3, 3).Font.Name = app.Theme.FontName)
    End With
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_actions__opt2"
    transcript = transcript & "|picks=" & gLastCommand & "/" & gCommandCount & "/" & _
        CStr(Not ShapeExists(host, "rdm_wid61_actions__opt1") _
            And host.Shapes("rdm_wid61_actions").TextFrame2.TextRange.Text = "Actions")
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_actions"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_actions__opt4"
    transcript = transcript & "|disabledRunsNothing=" & CStr(gCommandCount = 1)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_actions"
    transcript = transcript & "|alt=" & CStr(InStr(1, _
        host.Shapes("rdm_wid61_actions").AlternativeText, "Actions, menu button") > 0)

    app.MenuButton("actions").Focus
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|keys=" & gLastCommand & "/" & gCommandCount
    RdxReleaseKeys
    app.MenuButton("actions").Primary
    transcript = transcript & "|primary=" & CStr( _
        host.Shapes("rdm_wid61_actions").Fill.ForeColor.RGB = app.Theme.PrimaryColor _
        And host.Shapes("rdm_wid61_actions").TextFrame2.TextRange.Font.Fill.ForeColor.RGB _
            = app.Theme.OnPrimaryColor)

    ' A command wider than its button widens every row to fit it, and a
    ' press on the rows' far side, past the button, keeps the menu open.
    app.MenuButton("wide").AtRect(24, 200, 80, 26).Text("Wide") _
        .AddCommand("Export this view to a new sheet", "TestReDimWidgets.RecordCommand", _
            "Download") _
        .AddCommand "Copy", "TestReDimWidgets.RecordCommand", "Copy"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_wide"
    With host.Shapes("rdm_wid61_wide__opt1")
        transcript = transcript & "|fitsWords=" & CStr(.Width > 80 _
            And .TextFrame2.TextRange.BoundWidth + 12 <= .Width _
            And host.Shapes("rdm_wid61_wide__opt2").Width = .Width _
            And host.Shapes("rdm_wid61_wide__opt2").Left = .Left)
        farX = .Left + .Width - 4
        farY = .Top + .Height / 2
    End With
    ReDimUI.OverridePointer farX, farY
    ReDimUI.ForcePressEdge
    ReDimUI.PumpOnce
    transcript = transcript & "|farSideKeeps=" & _
        CStr(ShapeExists(host, "rdm_wid61_wide__opt1"))
    ReDimUI.ClearPointerOverride

    ' A long command paged into view widens the open menu to fit it, and
    ' an open menu follows its button when the button moves.
    app.MenuButton("many").AtRect(260, 200, 80, 26).Text "Many"
    For idx = 1 To 8
        app.MenuButton("many").AddCommand "Cmd " & idx, "TestReDimWidgets.RecordCommand"
    Next idx
    app.MenuButton("many").AddCommand "Export this view to a new sheet", _
        "TestReDimWidgets.RecordCommand"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_many"
    narrowWidth = host.Shapes("rdm_wid61_many__opt1").Width
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid61_many__optd"
    For idx = 1 To 8
        If ShapeExists(host, "rdm_wid61_many__opt" & idx) Then
            If InStr(host.Shapes("rdm_wid61_many__opt" & idx).TextFrame2.TextRange.Text, _
                "Export") > 0 Then longRow = "rdm_wid61_many__opt" & idx
        End If
    Next idx
    With host.Shapes(longRow)
        transcript = transcript & "|pagedWidens=" & CStr(narrowWidth = 80 And .Width > 80 _
            And .TextFrame2.TextRange.BoundWidth + 12 <= .Width)
    End With
    app.MenuButton("many").AtRect 200, 200, 80, 26
    With host.Shapes(longRow)
        transcript = transcript & "|followsFace=" & CStr(Abs(.Left - 200) < 0.5 _
            Or Abs(.Left + .Width - 280) < 0.5)
    End With
    ReDimUI.AutoPump True
    TestMenuButton = transcript
End Function

' Command palette: it lists the app's commands, then each button with a
' click handler, then each menu's commands. Typing filters it, and a pick
' closes it and runs the entry: a menu's command through the menu, a
' button as if clicked, an app command with the app as the sender.
' Leaving it cancels without running what was typed, and the palette key
' opens it on the app's sheet.
Public Function TestCommandPalette() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    gCommandCount = 0
    gLastCommand = vbNullString
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid62")
    app.AddCommand "Say hello", "TestReDimWidgets.RecordAppCommand", "Send"
    app.Button("save").AtRect(24, 24, 110, 30).Text("Save report") _
        .OnClick "TestReDimWidgets.RecordChange"
    app.Button("idle").AtRect(24, 70, 110, 30).Text "No handler"
    app.MenuButton("more").AtRect(150, 24, 100, 28).Text("More") _
        .AddCommand "Archive", "TestReDimWidgets.RecordCommand"
    app.CommandPalette
    app.Render

    app.OpenCommandPalette
    With app.ComboBox("mdl_pal_field")
        transcript = "lists=" & .ItemCount & ":" & .ItemTextAt(1) & "," & .ItemTextAt(2) & _
            "," & .ItemTextAt(3)
    End With
    transcript = transcript & "|focused=" & CStr( _
        ReDimUI.IsComponentFocused("wid62", "mdl_pal_field") _
        And host.Shapes("rdm_wid62_mdl_pal_card").Visible = msoTrue)
    TypeText "arch"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|menuEntry=" & gLastCommand & "/" & gCommandCount & "/" & _
        CStr(host.Shapes("rdm_wid62_mdl_pal_field").Visible = msoFalse)

    app.OpenCommandPalette
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|buttonEntry=" & gChangeCount

    app.OpenCommandPalette
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|appEntry=" & gLastCommand & "/" & gCommandCount

    app.OpenCommandPalette
    TypeText "Say"
    ReDimUI.EndKeyboardFocus
    transcript = transcript & "|leaveCancels=" & CStr(gCommandCount = 2 _
        And host.Shapes("rdm_wid62_mdl_pal_field").Visible = msoFalse)

    RdxOpenPalette
    transcript = transcript & "|keyOpens=" & CStr( _
        host.Shapes("rdm_wid62_mdl_pal_field").Visible = msoTrue)
    ReDimUI.EndKeyboardFocus
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestCommandPalette = transcript
End Function

' Table data tools: typing filters the rows on any cell and the footer
' names the filter; Backspace takes a letter back and Esc clears it.
' Ctrl+C copies every row shown, or the selected one, under the header,
' with dates the sheet reads back as dates. ExportTo writes the rows
' shown, sorted, to a range. EmptyText words an empty table, a filter
' that matches nothing says so, and a double click or Enter opens a row.
' Each paste yields first: Excel sees a clipboard written through Windows
' only once it has handled the change, which a user's Ctrl+V always allows.
Public Function TestTableData() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gCommandCount = 0
    gLastCommand = vbNullString
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid63")
    app.Table("orders").AtRect(24, 24, 360, 160).Columns "Item", "Qty", "Due"
    With app.Table("orders")
        .AddRow "Apple", 12, DateSerial(2026, 9, 1)
        .AddRow "Apricot", 3, DateSerial(2026, 9, 5)
        .AddRow "Banana", 7, DateSerial(2026, 9, 9)
        .AddRow "Cherry", 1, DateSerial(2026, 9, 12)
        .OnRowOpen "TestReDimWidgets.RecordRowOpen"
    End With
    app.Table("empty").AtRect(24, 220, 200, 80).Columns("Name").EmptyText "No orders yet"
    app.Render

    app.Table("orders").Focus
    RdxKeyChar "a"
    RdxKeyChar "p"
    transcript = "typed=" & app.Table("orders").ShownRowCount & "/" & _
        host.Shapes("rdm_wid63_orders__tf").TextFrame2.TextRange.Text
    RdxKeyChar "{BS}"
    transcript = transcript & "|backspace=" & app.Table("orders").ShownRowCount
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escClears=" & app.Table("orders").ShownRowCount & "/" & _
        CStr(ReDimUI.IsComponentFocused("wid63", "orders"))

    app.Table("orders").FilterRows "an"
    RdxKeyChar "{COPY}"
    PasteWhenReady host, host.Range("J1")
    transcript = transcript & "|copied=" & host.Range("J1").Value & "," & _
        host.Range("J2").Value & "," & host.Range("K2").Value & "," & _
        CStr(VarType(host.Range("L2").Value) = vbDate) & "," & CStr(IsEmpty(host.Range("J3").Value))
    app.Table("orders").FilterRows ""
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{COPY}"
    PasteWhenReady host, host.Range("J5")
    transcript = transcript & "|copiedRow=" & host.Range("J6").Value & "," & _
        CStr(IsEmpty(host.Range("J7").Value))
    app.Table("orders").FilterRows "an"
    RdxKeyChar "{COPY}"
    PasteWhenReady host, host.Range("J9")
    transcript = transcript & "|hiddenPick=" & host.Range("J10").Value
    RdxKeyChar "{ENTER}"
    transcript = transcript & "/" & gCommandCount

    app.Table("orders").SortBy 2
    app.Table("orders").FilterRows "r"
    app.Table("orders").ExportTo host.Range("N1")
    transcript = transcript & "|exported=" & host.Range("N1").Value & "," & _
        host.Range("N2").Value & "," & host.Range("N3").Value & "," & _
        CStr(IsEmpty(host.Range("N4").Value))

    app.Table("orders").FilterRows "zzz"
    transcript = transcript & "|noMatch=" & Trim$(Replace( _
        host.Shapes("rdm_wid63_orders__tr1").TextFrame2.TextRange.Text, vbTab, " "))
    transcript = transcript & "|emptyText=" & Trim$(Replace( _
        host.Shapes("rdm_wid63_empty__tr1").TextFrame2.TextRange.Text, vbTab, " "))

    app.Table("orders").FilterRows ""
    app.Table("orders").SortBy 0
    ReDimUI.DispatchShape "rdm_wid63_orders__tr3"
    ReDimUI.DispatchShape "rdm_wid63_orders__tr3"
    transcript = transcript & "|doubleOpens=" & gLastCommand & "/" & gCommandCount
    app.Table("orders").Focus
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterOpens=" & gCommandCount
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestTableData = transcript
End Function

' Pastes the clipboard once Excel can see what was just copied: a Win32
' clipboard write reaches Worksheet.Paste only after Excel processes the
' change, and another program can hold the clipboard for a moment, so the
' paste retries for up to a second before it raises the real error.
Private Sub PasteWhenReady(ByVal host As Worksheet, ByVal pasteAt As Range)
    Dim attemptNo As Long

    For attemptNo = 1 To 40
        DoEvents
        On Error Resume Next
        host.Paste pasteAt
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Sleep 25
    Next attemptNo
    host.Paste pasteAt
End Sub

' Masked fields: the face shows a dot for each character, focused or not,
' while InputValue and WritesTo keep the text; copy and cut take nothing.
Public Function TestMasked() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim faceNow As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid64")
    app.TextInput("plain").AtRect(24, 24, 160, 22).Text "sentinel"
    app.TextInput("pin").AtRect(24, 70, 160, 22).Masked.WritesTo "pinState"
    app.Render

    ReDimUI.DispatchShape "rdm_wid64_plain"
    RdxKeyChar "{SELECTALL}"
    RdxKeyChar "{COPY}"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid64_pin"
    TypeText "s3cret"
    faceNow = host.Shapes("rdm_wid64_pin").TextFrame2.TextRange.Text
    transcript = "focusedFace=" & CStr(InStr(faceNow, "s3cret") = 0 _
        And InStr(faceNow, String$(6, ChrW(8226))) > 0)
    RdxKeyChar "{SELECTALL}"
    RdxKeyChar "{COPY}"
    RdxKeyChar "{ENTER}"
    host.Paste host.Range("J1")
    transcript = transcript & "|copyTakesNothing=" & host.Range("J1").Value
    transcript = transcript & "|committed=" & app.State("pinState") & "/" & _
        app.TextInput("pin").InputValue
    transcript = transcript & "|restFace=" & CStr( _
        host.Shapes("rdm_wid64_pin").TextFrame2.TextRange.Text = String$(6, ChrW(8226)))
    transcript = transcript & "|alt=" & host.Shapes("rdm_wid64_pin").AlternativeText
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestMasked = transcript
End Function

' ValidateAll checks every enabled float field with Required or
' Validates the way a commit does, the ones on a tab not shown included,
' skips hidden and disabled fields, and focuses the first that fails in
' Tab order, turning its Tabs control to its tab first.
Public Function TestValidateAll() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid65")
    app.Tabs("tabs").AtRect(24, 24, 300, 30).Items("Name", "Contact").WritesTo "tabNow"
    app.TextInput("name").AtRect(24, 80, 200, 22).OnTab("tabs", 1).Text "Ada"
    app.TextInput("email").AtRect(24, 80, 200, 22).OnTab("tabs", 2).Required
    app.TextInput("mail2").AtRect(24, 130, 200, 22).OnTab("tabs", 2) _
        .Validates "TestReDimWidgets.CheckEmail"
    app.TextInput("gone").AtRect(260, 80, 160, 22).Required.Visible False
    app.TextInput("off").AtRect(260, 130, 160, 22).Required.Enabled False
    app.Render

    transcript = "first=" & CStr(app.ValidateAll) & "/" & app.Tabs("tabs").CurrentValue & _
        "/" & app.State("tabNow") & "/" & CStr(ReDimUI.IsComponentFocused("wid65", "email"))
    transcript = transcript & "|messages=" & app.TextInput("email").ValidationError & "," & _
        app.TextInput("mail2").ValidationError & "," & app.TextInput("gone").ValidationError & _
        "," & app.TextInput("off").ValidationError & "," & _
        CStr(host.Shapes("rdm_wid65_email").Visible = msoTrue)
    TypeText "ada@example.com"
    transcript = transcript & "|second=" & CStr(app.ValidateAll) & "/" & _
        CStr(ReDimUI.IsComponentFocused("wid65", "mail2")) & "/" & _
        app.TextInput("email").ValidationError
    TypeText "ada@home"
    transcript = transcript & "|third=" & CStr(app.ValidateAll) & "/" & _
        app.TextInput("mail2").ValidationError
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestValidateAll = transcript
End Function

' SetUIText replaces the words ReDim draws and announces on its own: a
' table's empty row and footer, a check list's select-all row, a
' required field's message, the Confirm buttons, and the descriptions
' screen readers get. Keys match in any case, a fill's own braces stay
' as typed, an unknown key raises, and ResetUIText restores English.
Public Function TestUIText() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim keyNames As Variant
    Dim rowNo As Long

    ReDimUI.AutoPump False
    ReDimUI.SetUIText "NoRows", "Keine Zeilen"
    ReDimUI.SetUIText "rowrange", "{0}-{1} von {2}"
    ReDimUI.SetUIText "FilterNote", "Filter ""{0}"": {1}"
    ReDimUI.SetUIText "SelectAll", "Alle ({0}/{1})"
    ReDimUI.SetUIText "Required", "Pflichtfeld"
    ReDimUI.SetUIText "OK", "Ja"
    ReDimUI.SetUIText "AltButton", "Knopf {0}"
    ReDimUI.SetUIText "AltSwitchOn", "Schalter, an"
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid66")
    app.Table("empty").AtRect(24, 24, 200, 60).Columns "Name"
    app.Table("many").AtRect(24, 100, 200, 120).Columns "Word"
    For rowNo = 1 To 6
        app.Table("many").AddRow "w" & rowNo
    Next rowNo
    app.CheckList("pick").AtRect(260, 24, 160, 90).Items "A", "B"
    app.TextInput("need").AtRect(260, 150, 160, 22).Required
    app.Button("go").AtRect(260, 200, 100, 30).Text "Go"
    app.Toggle("sw").AtRect(260, 250, 60, 24).Checked True
    app.Render

    transcript = "emptyRow=" & Trim$(Replace( _
        host.Shapes("rdm_wid66_empty__tr1").TextFrame2.TextRange.Text, vbTab, " "))
    transcript = transcript & "|footer=" & _
        host.Shapes("rdm_wid66_many__tf").TextFrame2.TextRange.Text
    app.Table("many").FilterRows "{1}"
    transcript = transcript & "|braces=" & _
        host.Shapes("rdm_wid66_many__tf").TextFrame2.TextRange.Text
    transcript = transcript & "|selectAll=" & _
        host.Shapes("rdm_wid66_pick__mt").TextFrame2.TextRange.Text
    ReDimUI.DispatchShape "rdm_wid66_need"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|required=" & app.TextInput("need").ValidationError
    transcript = transcript & "|alt=" & host.Shapes("rdm_wid66_go").AlternativeText & "/" & _
        host.Shapes("rdm_wid66_sw").AlternativeText
    transcript = transcript & "|readBack=" & ReDimUI.UIText("ROWRANGE")
    keyNames = ReDimUI.UITextKeys
    transcript = transcript & "|keys=" & (UBound(keyNames) - LBound(keyNames) + 1) & "/" & _
        keyNames(LBound(keyNames))
    On Error Resume Next
    ReDimUI.SetUIText "NoSuchText", "x"
    transcript = transcript & "|unknownRaises=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    app.Confirm "Title", "Message"
    transcript = transcript & "|confirm=" & _
        host.Shapes("rdm_wid66_mdl_ok").TextFrame2.TextRange.Text & "/" & _
        host.Shapes("rdm_wid66_mdl_cancel").TextFrame2.TextRange.Text
    app.CloseModal
    ReDimUI.ResetUIText
    app.Render
    transcript = transcript & "|reset=" & Trim$(Replace( _
        host.Shapes("rdm_wid66_empty__tr1").TextFrame2.TextRange.Text, vbTab, " ")) & "/" & _
        host.Shapes("rdm_wid66_go").AlternativeText
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestUIText = transcript
End Function

' A sparkline draws one polyline through its values, scaled to its
' rectangle less half a dot at each edge, with a dot on the last value.
' Blanks and text are left out, equal values run along the middle, one
' value is a dot alone, and screen readers hear a summary.
Public Function TestSparkline() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim lineShape As Shape
    Dim dotShape As Shape

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid67")
    host.Range("J1").Value = 1
    host.Range("J3").Value = 3
    app.Sparkline("trend").AtRect(24, 24, 125, 45) _
        .ValuesFrom Array(3, 5, "n/a", 4, Empty, 8, 6)
    app.Sparkline("flat").AtRect(24, 100, 100, 30).Danger.ValuesFrom Array(2, 2, 2)
    app.Sparkline("cells").AtRect(24, 150, 100, 30).ValuesFrom host.Range("J1:J3")
    app.Sparkline("none").AtRect 24, 200, 100, 30
    app.Render

    Set lineShape = host.Shapes("rdm_wid67_trend__sl")
    Set dotShape = host.Shapes("rdm_wid67_trend__sd")
    transcript = "nodes=" & lineShape.Nodes.Count
    transcript = transcript & "|lineBox=" & Round(lineShape.Left, 1) & "," & _
        Round(lineShape.Top, 1) & "," & Round(lineShape.Width, 1) & "," & _
        Round(lineShape.Height, 1)
    transcript = transcript & "|dotCenter=" & Round(dotShape.Left + dotShape.Width / 2, 1) & _
        "," & Round(dotShape.Top + dotShape.Height / 2, 1)
    transcript = transcript & "|ink=" & CStr(lineShape.Line.ForeColor.RGB = app.Theme.PrimaryColor _
        And dotShape.Fill.ForeColor.RGB = app.Theme.PrimaryColor And _
        lineShape.Fill.Visible = msoFalse)
    transcript = transcript & "|flat=" & Round(host.Shapes("rdm_wid67_flat__sl").Height, 1) & _
        "/" & CStr(host.Shapes("rdm_wid67_flat__sl").Line.ForeColor.RGB = app.Theme.DangerColor)
    transcript = transcript & "|alt=" & host.Shapes("rdm_wid67_trend").AlternativeText & ";" & _
        host.Shapes("rdm_wid67_cells").AlternativeText & ";" & _
        host.Shapes("rdm_wid67_none").AlternativeText
    transcript = transcript & "|emptyParts=" & CStr(ShapeExists(host, "rdm_wid67_none__sl") _
        Or ShapeExists(host, "rdm_wid67_none__sd"))

    app.Sparkline("trend").ValuesFrom Array(7)
    transcript = transcript & "|single=" & CStr(ShapeExists(host, "rdm_wid67_trend__sl")) & _
        "/" & Round(host.Shapes("rdm_wid67_trend__sd").Left + 2.5, 1)
    app.Sparkline("trend").Visible False
    transcript = transcript & "|hidden=" & _
        CStr(host.Shapes("rdm_wid67_trend__sd").Visible = msoFalse)
    app.Sparkline("trend").Remove
    transcript = transcript & "|removed=" & CStr(ShapeExists(host, "rdm_wid67_trend__sd"))
    ReDimUI.AutoPump True
    TestSparkline = transcript
End Function

' A flip of a switch drawn in place glides its knob across over the
' pump's frames instead of jumping, and a tab strip's bar glides to the
' tab shown; with motion reduced both land at once.
Public Function TestGlide() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim knob As Shape
    Dim bar As Shape
    Dim tabThree As Shape
    Dim startLeft As Double
    Dim frameNo As Long

    ReDimUI.AutoPump False
    ' Full motion whatever the machine's Windows animation setting.
    ReDimUI.ReduceMotion False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid68")
    app.Toggle("sw").AtRect 24, 24, 44, 22
    app.Tabs("tabs").AtRect(24, 70, 300, 30).Items "One", "Two", "Three"
    app.Render

    Set knob = host.Shapes("rdm_wid68_sw__knob")
    startLeft = knob.Left
    ReDimUI.DispatchShape "rdm_wid68_sw"
    transcript = "start=" & Round(startLeft, 2) & "/" & CStr(Abs(knob.Left - startLeft) < 0.01)
    ReDimUI.PumpOnce
    transcript = transcript & "|moving=" & CStr(knob.Left > startLeft + 1 And knob.Left < 47)
    For frameNo = 1 To 10
        ReDimUI.PumpOnce
    Next frameNo
    transcript = transcript & "|lands=" & Round(knob.Left, 2)

    Set bar = host.Shapes("rdm_wid68_tabs__ti")
    Set tabThree = host.Shapes("rdm_wid68_tabs__tb3")
    startLeft = bar.Left
    ReDimUI.DispatchShape "rdm_wid68_tabs__tb3"
    transcript = transcript & "|barStays=" & CStr(Abs(bar.Left - startLeft) < 0.01)
    For frameNo = 1 To 12
        ReDimUI.PumpOnce
    Next frameNo
    transcript = transcript & "|barLands=" & CStr(bar.Left > tabThree.Left _
        And bar.Left + bar.Width < tabThree.Left + tabThree.Width)
    transcript = transcript & "|settled=" & CStr(Not ReDimUI.App("wid68").HasAppWork)

    ReDimUI.ReduceMotion True
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid68_sw"
    transcript = transcript & "|reducedJumps=" & Round(knob.Left, 2)
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestGlide = transcript
End Function

' StateKeys lists every state key once, in the order it was first set,
' whether code or a control's WritesTo set it; an empty store lists none.
Public Function TestStateKeys() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim keyList As Variant

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid69")
    keyList = app.StateKeys
    transcript = "empty=" & (UBound(keyList) - LBound(keyList) + 1)
    app.SetState "user", "ada"
    app.SetState "count", 1
    app.SetState "user", "grace"
    app.Toggle("dark").AtRect(24, 24, 44, 22).WritesTo "darkMode"
    app.Render
    ReDimUI.DispatchShape "rdm_wid69_dark"
    keyList = app.StateKeys
    transcript = transcript & "|keys=" & Join(keyList, ",") & "|from=" & LBound(keyList)
    ReDimUI.AutoPump True
    TestStateKeys = transcript
End Function

' A Toggle, TickBox, or Expander bound with BindValue follows True and
' False from the store, and with WritesTo on the same key a click and the
' store stay in step.
Public Function TestBoolBinding() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid70")
    app.SetState "on", False
    app.Toggle("sw").AtRect(24, 24, 44, 22).WritesTo("on").BindValue "on"
    app.TickBox("tick").AtRect(24, 60, 120, 18).Text("Tick").BindValue "on"
    app.Expander("more").AtRect(24, 90, 200, 28).Text("More").BindValue "on"
    app.Render
    transcript = "start=" & CStr(app.Toggle("sw").IsChecked) & "/" & _
        CStr(app.TickBox("tick").IsChecked) & "/" & CStr(app.Expander("more").IsExpanded)
    app.SetState "on", True
    transcript = transcript & "|follows=" & CStr(app.Toggle("sw").IsChecked) & "/" & _
        CStr(app.TickBox("tick").IsChecked) & "/" & CStr(app.Expander("more").IsExpanded)
    ReDimUI.DispatchShape "rdm_wid70_sw"
    transcript = transcript & "|clickWrites=" & CStr(app.State("on")) & "/" & _
        CStr(app.TickBox("tick").IsChecked)
    ReDimUI.AutoPump True
    TestBoolBinding = transcript
End Function

' A press read as off an open list closes it once the button is back up
' and the grace for a click has passed. A click Excel delivers on the
' list's own pager in the meantime keeps it open, as happens when the
' pointer read lands just off the list, and a press off with no click
' still closes it.
Public Function TestPressOffRelease() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    ReDimUI.OverridePointer 900, 5, False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid71")
    app.SelectBox("pick").AtRect(24, 24, 140, 22).Items "A", "B", "C", "D", "E", "F", _
        "G", "H", "I", "J", "K", "L"
    app.Render
    ReDimUI.DispatchShape "rdm_wid71_pick"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid71_pick__optd"

    ReDimUI.OverridePointer 600, 400, True
    ReDimUI.PumpOnce
    transcript = "held=" & CStr(ShapeExists(host, "rdm_wid71_pick__opt1"))
    ReDimUI.OverridePointer 600, 400, False
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid71_pick__optu"
    ReDimUI.PumpOnce
    Sleep 150
    ReDimUI.PumpOnce
    transcript = transcript & "|clickKeeps=" & CStr(ShapeExists(host, "rdm_wid71_pick__opt1")) & _
        "/" & Trim$(Mid$(host.Shapes("rdm_wid71_pick__opt1").TextFrame2.TextRange.Text, 2))

    ReDimUI.OverridePointer 600, 400, True
    ReDimUI.PumpOnce
    ReDimUI.OverridePointer 600, 400, False
    ReDimUI.PumpOnce
    transcript = transcript & "|waits=" & CStr(ShapeExists(host, "rdm_wid71_pick__opt1"))
    Sleep 150
    ReDimUI.PumpOnce
    transcript = transcript & "|closes=" & CStr(Not ShapeExists(host, "rdm_wid71_pick__opt1"))
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestPressOffRelease = transcript
End Function

' Clicking through a list's pages at one spot (issue #2): a list that
' pages keeps both pager rows in place, so the down pager stays under the
' pointer from the first page to the last, where it turns inert instead
' of leaving an item row to be picked. The up pager at the top is inert
' the same way.
Public Function TestPagerHoldsPlace() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim names(0 To 19) As String
    Dim idx As Long
    Dim pagerTop As Double
    Dim steady As Boolean

    For idx = 0 To 19
        names(idx) = "Item" & Format$(idx + 1, "00")
    Next idx
    ReDimUI.AutoPump False
    ReDimUI.OverridePointer 900, 5, False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid75")
    app.SelectBox("pick").AtRect(24, 24, 140, 22).ItemsFrom names
    app.Render
    ReDimUI.DispatchShape "rdm_wid75_pick"
    transcript = "upInertAtTop=" & CStr(PagerInert(host, "rdm_wid75_pick__optu"))
    pagerTop = host.Shapes("rdm_wid75_pick__optd").Top
    steady = True
    For idx = 1 To 4
        Sleep 200
        ReDimUI.OverridePointer 90, pagerTop + 11, True
        ReDimUI.PumpOnce
        ReDimUI.OverridePointer 90, pagerTop + 11, False
        ReDimUI.PumpOnce
        ReDimUI.DispatchShape "rdm_wid75_pick__optd"
        Sleep 150
        ReDimUI.PumpOnce
        If Not ShapeExists(host, "rdm_wid75_pick__optd") Then
            steady = False
        ElseIf host.Shapes("rdm_wid75_pick__optd").Top <> pagerTop Then
            steady = False
        End If
    Next idx
    transcript = transcript & "|downStays=" & CStr(steady)
    transcript = transcript & "|stillOpen=" & CStr(app.SelectBox("pick").ListIsOpen)
    transcript = transcript & "|nothingPicked=" & CStr(app.SelectBox("pick").CurrentValue = 0)
    transcript = transcript & "|downInertAtEnd=" & _
        CStr(PagerInert(host, "rdm_wid75_pick__optd"))
    transcript = transcript & "|lastRow=" & _
        Trim$(Mid$(host.Shapes("rdm_wid75_pick__opt8").TextFrame2.TextRange.Text, 2))
    ReDimUI.ClearPointerOverride
    ReDimUI.AutoPump True
    TestPagerHoldsPlace = transcript
End Function

' A button or arrow with nowhere to go reads muted: a stepper's minus at
' its minimum and plus at its maximum, and a transfer panel's up arrow on
' its first page and down arrow on its last.
Public Function TestDeadEnds() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim names(0 To 11) As String
    Dim idx As Long
    Dim live As Long
    Dim muted As Long

    For idx = 0 To 11
        names(idx) = "Item" & Format$(idx + 1, "00")
    Next idx
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid76")
    app.Stepper("qty").AtRect(24, 24, 120, 24).SliderRange(0, 3, 1).Value 0
    app.TransferList("pool").AtRect(24, 80, 380, 128).ItemsFrom names
    app.Render
    live = app.Theme.OnSurfaceColor
    muted = app.Theme.OnMutedColor
    transcript = "minAtMin=" & CStr(InkOf(host, "rdm_wid76_qty__minus") = muted And _
        InkOf(host, "rdm_wid76_qty__plus") = live)
    app.Stepper("qty").Value 1
    transcript = transcript & "|bothLive=" & CStr(InkOf(host, "rdm_wid76_qty__minus") = live And _
        InkOf(host, "rdm_wid76_qty__plus") = live)
    app.Stepper("qty").Value 3
    transcript = transcript & "|plusAtMax=" & CStr(InkOf(host, "rdm_wid76_qty__minus") = live And _
        InkOf(host, "rdm_wid76_qty__plus") = muted)
    transcript = transcript & "|firstPage=" & CStr(InkOf(host, "rdm_wid76_pool__alu") = muted And _
        InkOf(host, "rdm_wid76_pool__ald") = live)
    For idx = 1 To 3
        Sleep 200
        ReDimUI.DispatchShape "rdm_wid76_pool__ald"
    Next idx
    transcript = transcript & "|lastPage=" & CStr(InkOf(host, "rdm_wid76_pool__alu") = live And _
        InkOf(host, "rdm_wid76_pool__ald") = muted)
    ' Move buttons with nothing to move read muted.
    transcript = transcript & "|movesIdle=" & CStr(InkOf(host, "rdm_wid76_pool__mvr") = muted And _
        InkOf(host, "rdm_wid76_pool__mvar") = live And _
        InkOf(host, "rdm_wid76_pool__mvl") = muted And _
        InkOf(host, "rdm_wid76_pool__mval") = muted)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid76_pool__al1"
    transcript = transcript & "|selectionWakesMove=" & CStr(InkOf(host, "rdm_wid76_pool__mvr") = live)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid76_pool__mvr"
    transcript = transcript & "|chosenWakesBack=" & CStr(InkOf(host, "rdm_wid76_pool__mval") = live _
        And InkOf(host, "rdm_wid76_pool__mvr") = muted)
    On Error Resume Next
    app.Stepper("qty").SliderRange 0, 3, 0
    transcript = transcript & "|zeroStepRefused=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    ReDimUI.AutoPump True
    TestDeadEnds = transcript
End Function

' A RadioGroup or CheckList emptied of its items takes its rows down,
' select-all header included, and draws them again when items return.
Public Function TestEmptiedLists() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid77")
    app.RadioGroup("tier").AtRect(24, 24, 140, 60).Items "Low", "Mid", "High"
    app.CheckList("tags").AtRect(200, 24, 170, 100).Items "Red", "Green", "Blue"
    app.Render
    app.RadioGroup("tier").ClearItems
    app.CheckList("tags").ClearItems
    transcript = "radioCleared=" & CStr(Not ShapeExists(host, "rdm_wid77_tier__c2") And _
        Not ShapeExists(host, "rdm_wid77_tier__t1") And _
        host.Shapes("rdm_wid77_tier").Visible = msoFalse)
    transcript = transcript & "|checkCleared=" & CStr(Not ShapeExists(host, "rdm_wid77_tags__b2") And _
        Not ShapeExists(host, "rdm_wid77_tags__t1") And _
        Not ShapeExists(host, "rdm_wid77_tags__mt") And _
        host.Shapes("rdm_wid77_tags").Visible = msoFalse)
    app.RadioGroup("tier").AddItem "Solo"
    app.CheckList("tags").AddItem "Gold"
    transcript = transcript & "|radioReturns=" & CStr(host.Shapes("rdm_wid77_tier").Visible = msoTrue And _
        host.Shapes("rdm_wid77_tier__t1").TextFrame2.TextRange.Text = "Solo")
    transcript = transcript & "|checkReturns=" & CStr(host.Shapes("rdm_wid77_tags").Visible = msoTrue And _
        host.Shapes("rdm_wid77_tags__t1").TextFrame2.TextRange.Text = "Gold" And _
        ShapeExists(host, "rdm_wid77_tags__mt"))
    ReDimUI.AutoPump True
    TestEmptiedLists = transcript
End Function

' Ctrl+A in a focused CheckList checks every row and fires OnChange
' once; again, it clears nothing. In a TransferList it selects every row
' of the cursor's panel, which the move button then carries across.
Public Function TestSelectAllKeys() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid78")
    app.CheckList("tags").AtRect(24, 24, 170, 100).Items("Red", "Green", "Blue", "Gold") _
        .WritesTo("tagState").OnChange "TestReDimWidgets.RecordChange"
    app.TransferList("pool").AtRect(24, 150, 380, 128).Items "A", "B", "C", "D", "E"
    app.Render
    app.CheckList("tags").Focus
    RdxKeyChar "{SELECTALL}"
    transcript = "checksAll=" & app.CheckList("tags").CheckedCount & "/" & gChangeCount
    RdxKeyChar "{SELECTALL}"
    transcript = transcript & "|againKeeps=" & app.CheckList("tags").CheckedCount & "/" & gChangeCount
    app.TransferList("pool").Focus
    RdxKeyChar "{SELECTALL}"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid78_pool__mvr"
    transcript = transcript & "|transferSelectsAll=" & app.TransferList("pool").ChosenCount
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestSelectAllKeys = transcript
End Function

' SelectedText and SelectedValue read the selection the same way on every
' item control: the item, its value, or empty while nothing is selected,
' never a placeholder or a combo's free text.
Public Function TestSelectionReaders() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid80")
    app.SelectBox("size").AtRect(24, 24, 140, 22).Text("Pick a size").Items "S", "M", "L"
    app.RadioGroup("tier").AtRect 24, 60, 140, 60
    app.RadioGroup("tier").AddItem "Low", , 10
    app.RadioGroup("tier").AddItem "High", , 20
    app.Tabs("pages").AtRect(200, 24, 240, 26).Items "One", "Two"
    app.ComboBox("find").AtRect(200, 70, 150, 22).Items "Alpha", "Beta"
    app.Render
    transcript = "noneSelected=" & CStr(app.SelectBox("size").SelectedText = "" And _
        app.SelectBox("size").CurrentText = "Pick a size" And _
        IsEmpty(app.SelectBox("size").SelectedValue))
    app.SelectBox("size").Value 2
    transcript = transcript & "|select=" & app.SelectBox("size").SelectedText & "/" & _
        app.SelectBox("size").SelectedValue
    app.RadioGroup("tier").Value 2
    transcript = transcript & "|radio=" & app.RadioGroup("tier").SelectedText & "/" & _
        app.RadioGroup("tier").SelectedValue
    transcript = transcript & "|tabs=" & app.Tabs("pages").SelectedText
    With app.ComboBox("find")
        .InputValue = "Beta"
        transcript = transcript & "|comboItem=" & .SelectedText
        .InputValue = "Gamma"
        transcript = transcript & "|comboFree=[" & .SelectedText & "]"
    End With
    On Error Resume Next
    transcript = transcript & app.Label("note").SelectedText
    transcript = transcript & "|otherKindRefused=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    ReDimUI.AutoPump True
    TestSelectionReaders = transcript
End Function

' A CheckList or RadioGroup whose rows would squeeze under 18 points
' shows a window of rows with paging arrows at its right edge: the arrows
' page it (muted at the ends, repeating while held), rows keep their
' item numbers for clicks, and the keys move the window with the cursor
' or selection. A list whose rows fit shows every row and no arrows.
Public Function TestRowListWindow() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim names(0 To 19) As String
    Dim idx As Long

    For idx = 0 To 19
        names(idx) = "Item" & Format$(idx + 1, "00")
    Next idx
    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid84")
    app.CheckList("tags").AtRect(24, 24, 170, 100).ItemsFrom(names) _
        .OnChange "TestReDimWidgets.RecordChange"
    app.RadioGroup("tier").AtRect(220, 24, 150, 60).ItemsFrom names
    app.CheckList("short").AtRect(400, 24, 170, 100).Items "A", "B", "C"
    app.Render
    ' 100 points with a select-all header: 18-point rows fit four items.
    transcript = "checkWindow=" & CStr(IsShown(host, "rdm_wid84_tags__t4") And _
        Not IsShown(host, "rdm_wid84_tags__t5") And _
        host.Shapes("rdm_wid84_tags__t2").Height >= 18)
    transcript = transcript & "|arrowsAtTop=" & CStr( _
        InkOf(host, "rdm_wid84_tags__lu") = app.Theme.OnMutedColor And _
        InkOf(host, "rdm_wid84_tags__ld") = app.Theme.OnSurfaceColor)
    transcript = transcript & "|shortHasNone=" & CStr(Not ShapeExists(host, "rdm_wid84_short__ld") _
        And IsShown(host, "rdm_wid84_short__t3"))
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid84_tags__ld"
    transcript = transcript & "|arrowPages=" & CStr(Not IsShown(host, "rdm_wid84_tags__t1") And _
        IsShown(host, "rdm_wid84_tags__t4") And IsShown(host, "rdm_wid84_tags__t7") And _
        host.Shapes("rdm_wid84_tags__t4").Top < host.Shapes("rdm_wid84_tags__t5").Top)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid84_tags__t6"
    transcript = transcript & "|rowClickMaps=" & CStr(app.CheckList("tags").IsItemChecked(6)) & _
        "/" & gChangeCount
    HoldOnPart host, "rdm_wid84_tags__ld"
    transcript = transcript & "|holdPages=" & CStr(Not IsShown(host, "rdm_wid84_tags__t7"))
    app.CheckList("tags").Focus
    RdxKeyChar "{HOME}"
    transcript = transcript & "|homeTop=" & CStr(IsShown(host, "rdm_wid84_tags__t1"))
    RdxKeyChar "{END}"
    transcript = transcript & "|endBottom=" & CStr(IsShown(host, "rdm_wid84_tags__t20") And _
        InkOf(host, "rdm_wid84_tags__ld") = app.Theme.OnMutedColor)
    RdxKeyChar "{HOME}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{PGDN}"
    transcript = transcript & "|pageDownKey=" & CStr(IsShown(host, "rdm_wid84_tags__t5") And _
        Not IsShown(host, "rdm_wid84_tags__t1"))
    app.RadioGroup("tier").Focus
    RdxKeyChar "{HOME}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|radioFollows=" & app.RadioGroup("tier").CurrentValue & "/" & _
        CStr(IsShown(host, "rdm_wid84_tier__t4") And Not IsShown(host, "rdm_wid84_tier__t1"))
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestRowListWindow = transcript
End Function

' A Filterable SelectBox: a click opens it with the keys, letters filter
' the list and show on the face, the arrows and Enter work among the
' matches, a click on a filtered row takes that row's item, Backspace
' and Esc undo the filter, and a filter that matches nothing says so.
Public Function TestFilterableSelect() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid85")
    app.SelectBox("fruit").AtRect(24, 24, 150, 22).Text("Pick a fruit").Items("Apple", _
        "Apricot", "Banana", "Cherry", "Mango", "Orange", "Grape", "Lemon", "Lime", _
        "Peach", "Pear", "Plum").Filterable
    app.Render
    ReDimUI.DispatchShape "rdm_wid85_fruit"
    transcript = "clickTakesKeys=" & CStr(ReDimUI.FocusedComponentId = "fruit")
    RdxKeyChar "a"
    RdxKeyChar "n"
    transcript = transcript & "|filtered=" & RowItem(host, "rdm_wid85_fruit__opt1") & "," & _
        RowItem(host, "rdm_wid85_fruit__opt3") & "/" & _
        CStr(Not ShapeExists(host, "rdm_wid85_fruit__opt4")) & "/" & _
        host.Shapes("rdm_wid85_fruit").TextFrame2.TextRange.Text
    RdxKeyChar "{DOWN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|enterTakes=" & app.SelectBox("fruit").SelectedText & "/" & _
        CStr(Not app.SelectBox("fruit").ListIsOpen)
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid85_fruit"
    RdxKeyChar "p"
    RdxKeyChar "e"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid85_fruit__opt3"
    transcript = transcript & "|clickTakesMatch=" & app.SelectBox("fruit").SelectedText
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid85_fruit"
    RdxKeyChar "z"
    RdxKeyChar "z"
    transcript = transcript & "|noMatches=" & CStr(ShapeExists(host, "rdm_wid85_fruit__optn")) & _
        "/" & Replace(host.Shapes("rdm_wid85_fruit__optn").TextFrame2.TextRange.Text, vbTab, "")
    RdxKeyChar "{BS}"
    RdxKeyChar "{BS}"
    transcript = transcript & "|backspaceRestores=" & CStr(ShapeExists(host, "rdm_wid85_fruit__opt8"))
    RdxKeyChar "l"
    RdxKeyChar "{ESC}"
    transcript = transcript & "|escClearsFirst=" & CStr(app.SelectBox("fruit").ListIsOpen And _
        host.Shapes("rdm_wid85_fruit").TextFrame2.TextRange.Text = "Pear")
    RdxKeyChar "{ESC}"
    On Error Resume Next
    app.Button("go").Filterable
    transcript = transcript & "|otherKindRefused=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFilterableSelect = transcript
End Function

' The Windows edit chords and Ctrl+Page keys are captured, Ctrl+Page
' Down steps a tab strip and a calendar's year, and a letter from the
' keyboard follows Caps Lock.
Public Function TestEditChords() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim capsOn As Boolean

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid86")
    app.Tabs("pages").AtRect(24, 24, 240, 26).Items "One", "Two", "Three"
    app.DatePicker("due").AtRect(24, 70, 150, 24).PickDate DateSerial(2026, 3, 10)
    app.TextInput("name").AtRect 24, 110, 150, 22
    app.Stepper("qty").AtRect(24, 150, 120, 24).SliderRange(0, 500, 1).Value 5
    app.Render
    transcript = "bound=" & CStr(InStr(RdxCapturedCodes("{BS}"), Chr$(1) & "+{BS}" & Chr$(1)) > 0 _
        And InStr(RdxCapturedCodes("{CUT}"), "+{DEL}") > 0 _
        And InStr(RdxCapturedCodes("{PASTE}"), "+{INSERT}") > 0 _
        And InStr(RdxCapturedCodes("{COPY}"), "^{INSERT}") > 0 _
        And InStr(RdxCapturedCodes("{CTRLPGDN}"), "^{PGDN}") > 0)
    app.Tabs("pages").Focus
    RdxKeyChar "{CTRLPGDN}"
    transcript = transcript & "|tabsStep=" & app.Tabs("pages").SelectedText
    app.DatePicker("due").Focus
    RdxKeyChar "{ALTDOWN}"
    RdxKeyChar "{CTRLPGDN}"
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|calendarYear=" & Format$(app.DatePicker("due").PickedDate, "yyyy-mm-dd")
    capsOn = (GetKeyState(&H14) And 1) <> 0
    app.TextInput("name").Focus
    RdxKeyLetter "a"
    RdxKeyLetter "B"
    transcript = transcript & "|capsFollowed=" & CStr(app.TextInput("name").InputValue = _
        IIf(capsOn, "Ab", "aB"))
    ' Digits typed into a focused stepper set its value, clamped.
    app.Stepper("qty").Focus
    RdxKeyChar "2"
    RdxKeyChar "5"
    RdxKeyChar "0"
    transcript = transcript & "|stepperTyped=" & app.Stepper("qty").CurrentValue
    Sleep 1100
    RdxKeyChar "9"
    RdxKeyChar "9"
    RdxKeyChar "9"
    transcript = transcript & "|stepperClamped=" & app.Stepper("qty").CurrentValue
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestEditChords = transcript
End Function

' A one-line face at rest cuts a long text with an ellipsis on one line,
' keeps the whole text as its value, drops the ellipsis while focused,
' fits again when it leaves, and shows a short text whole.
Public Function TestFaceEllipsis() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim faceWords As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid87")
    app.TextInput("mail").AtRect 24, 24, 120, 22
    app.SelectBox("pick").AtRect(24, 70, 110, 24).Items("A very long selection", "Short").Value 1
    app.Render
    With app.TextInput("mail")
        .InputValue = "someone.with.a.long.name@example.com"
    End With
    faceWords = host.Shapes("rdm_wid87_mail").TextFrame2.TextRange.Text
    transcript = "fieldCut=" & CStr(Right$(faceWords, 1) = ChrW(8230) And _
        Left$(faceWords, 8) = "someone." And _
        app.TextInput("mail").InputValue = "someone.with.a.long.name@example.com")
    faceWords = host.Shapes("rdm_wid87_pick").TextFrame2.TextRange.Text
    transcript = transcript & "|selectCut=" & CStr(Right$(faceWords, 1) = ChrW(8230) And _
        app.SelectBox("pick").SelectedText = "A very long selection")
    ' Focused, the field shows its tail after a leading ellipsis instead.
    app.TextInput("mail").Focus
    faceWords = host.Shapes("rdm_wid87_mail").TextFrame2.TextRange.Text
    transcript = transcript & "|focusedTail=" & CStr(Left$(faceWords, 1) = ChrW(8230) And _
        InStr(faceWords, "@example.com") > 0)
    RdxKeyChar "{TAB}"
    RdxReleaseKeys
    transcript = transcript & "|blurCutsAgain=" & CStr(Right$(host.Shapes("rdm_wid87_mail") _
        .TextFrame2.TextRange.Text, 1) = ChrW(8230))
    app.SelectBox("pick").Value 2
    transcript = transcript & "|shortWhole=" & host.Shapes("rdm_wid87_pick").TextFrame2.TextRange.Text
    ReDimUI.AutoPump True
    TestFaceEllipsis = transcript
End Function

' True when a part exists and is visible.
Private Function IsShown(ByVal host As Worksheet, ByVal shapeName As String) As Boolean
    If Not ShapeExists(host, shapeName) Then Exit Function
    IsShown = (host.Shapes(shapeName).Visible = msoTrue)
End Function

' A Toggle's Text draws as a caption right of the switch, sized to its
' words; a click on it flips the switch, the alternative text leads with
' it, and Text "" takes it away.
Public Function TestToggleCaption() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    ReDimUI.ReduceMotion True
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid83")
    app.Toggle("dark").AtRect(24, 24, 44, 22).Text("Dark mode").WritesTo "darkMode"
    app.Toggle("bare").AtRect 24, 60, 44, 22
    app.Stack("row").AtRect(24, 120, 0, 0).Across.Gap 10
    app.Toggle("inRow").Sized(44, 22).Text("Wide caption here").InStack "row"
    app.Label("after").Sized(60, 18).Text("Next").InStack "row"
    app.Render
    With host.Shapes("rdm_wid83_dark__lbl")
        transcript = "captionRight=" & CStr(.Left >= 24 + 44 And _
            .TextFrame2.TextRange.Text = "Dark mode" And .Width < 120)
    End With
    With host.Shapes("rdm_wid83_inRow__lbl")
        transcript = transcript & "|stackClears=" & CStr(host.Shapes("rdm_wid83_after").Left _
            >= .Left + .Width + 9)
    End With
    transcript = transcript & "|bareHasNone=" & CStr(Not ShapeExists(host, "rdm_wid83_bare__lbl"))
    ReDimUI.DispatchShape "rdm_wid83_dark__lbl"
    transcript = transcript & "|captionFlips=" & CStr(app.State("darkMode"))
    transcript = transcript & "|altLeads=" & host.Shapes("rdm_wid83_dark").AlternativeText
    app.Toggle("dark").Text ""
    transcript = transcript & "|textGone=" & CStr(Not ShapeExists(host, "rdm_wid83_dark__lbl"))
    ReDimUI.ReduceMotion
    ReDimUI.AutoPump True
    TestToggleCaption = transcript
End Function

' Required on a SelectBox, DatePicker, and RadioGroup: ValidateAll fails
' while nothing is picked, shows the message under each, and puts focus
' on the first; a pick takes its message down, and ValidateAll passes
' once all three hold a pick.
Public Function TestRequiredPicks() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid82")
    app.SelectBox("size").AtRect(24, 24, 140, 22).Text("Pick a size").Items("S", "M") _
        .Required
    app.DatePicker("due").AtRect(24, 70, 140, 22).Required True, "Pick a due date"
    app.RadioGroup("tier").AtRect(24, 116, 140, 40).Items("Low", "High").Required
    app.Stack("form").AtRect(300, 24, 200, 0).Gap 6
    app.SelectBox("kind").Sized(150, 22).Items("A", "B").Required.InStack "form"
    app.Label("after").Sized(150, 18).Text("Next").InStack "form"
    app.Render
    transcript = "stackRoom=" & CStr(host.Shapes("rdm_wid82_after").Top >= _
        host.Shapes("rdm_wid82_kind").Top + 22 + 6 + 12)
    transcript = transcript & "|fails=" & CStr(Not app.ValidateAll)
    transcript = transcript & "|messages=" & _
        host.Shapes("rdm_wid82_size__me").TextFrame2.TextRange.Text & "/" & _
        host.Shapes("rdm_wid82_due__me").TextFrame2.TextRange.Text & "/" & _
        CStr(ShapeExists(host, "rdm_wid82_tier__me"))
    transcript = transcript & "|firstFocused=" & ReDimUI.FocusedComponentId
    RdxReleaseKeys
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid82_size"
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid82_size__opt2"
    transcript = transcript & "|pickClears=" & CStr(Not ShapeExists(host, "rdm_wid82_size__me"))
    app.DatePicker("due").PickDate DateSerial(2026, 10, 1)
    app.RadioGroup("tier").Value 1
    app.SelectBox("kind").Value 1
    transcript = transcript & "|passes=" & CStr(app.ValidateAll)
    transcript = transcript & "|allClear=" & CStr(Not ShapeExists(host, "rdm_wid82_due__me") _
        And Not ShapeExists(host, "rdm_wid82_tier__me"))
    ' ErrorText speaks under a pick whether it is required or not.
    app.SelectBox("kind").Required(False).ErrorText "That kind is sold out"
    transcript = transcript & "|errorTextShows=" & _
        host.Shapes("rdm_wid82_kind__me").TextFrame2.TextRange.Text
    app.SelectBox("kind").ErrorText ""
    transcript = transcript & "|errorTextClears=" & CStr(Not ShapeExists(host, "rdm_wid82_kind__me"))
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestRequiredPicks = transcript
End Function

' ItemEnabled on a RadioGroup and a ComboBox: a disabled item reads
' muted, a click or the keys never take it, and a combo's highlight and
' suggestion pass over it.
Public Function TestItemEnabledKinds() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gChangeCount = 0
    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid81")
    app.RadioGroup("tier").AtRect(24, 24, 140, 60).Items("Low", "Mid", "High") _
        .Value(1).OnChange "TestReDimWidgets.RecordChange"
    app.RadioGroup("tier").ItemEnabled 2, False
    app.ComboBox("find").AtRect(200, 24, 150, 22).Items "Apple", "Apricot", "Banana"
    app.ComboBox("find").ItemEnabled 1, False
    app.Render
    transcript = "radioMuted=" & CStr(InkOf(host, "rdm_wid81_tier__t2") = app.Theme.OnMutedColor)
    ReDimUI.DispatchShape "rdm_wid81_tier__t2"
    transcript = transcript & "|radioClickRefused=" & app.RadioGroup("tier").CurrentValue & "/" & gChangeCount
    app.RadioGroup("tier").Focus
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|radioKeySkips=" & app.RadioGroup("tier").CurrentValue
    RdxKeyChar "{UP}"
    transcript = transcript & "|radioKeyBack=" & app.RadioGroup("tier").CurrentValue

    app.ComboBox("find").Focus
    RdxKeyChar "{DOWN}"
    transcript = transcript & "|comboHighlightSkips=" & _
        CStr(host.Shapes("rdm_wid81_find__opt2").Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    RdxKeyChar "{ENTER}"
    transcript = transcript & "|comboTakes=" & app.ComboBox("find").SelectedText
    Sleep 200
    ReDimUI.DispatchShape "rdm_wid81_find"
    BackspaceAll app.ComboBox("find")
    RdxKeyChar "A"
    RdxKeyChar "p"
    RdxKeyChar "{TAB}"
    transcript = transcript & "|suggestionSkips=" & app.ComboBox("find").InputValue
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestItemEnabledKinds = transcript
End Function

' A long face text stops short of the drop caret on a SelectBox, a
' DatePicker, and a ComboBox, measured from where the text is drawn.
Public Function TestFaceClearsCaret() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String
    Dim faceName As Variant

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "wid79")
    app.SelectBox("pick").AtRect(24, 24, 110, 24).Items("Wwwwwwwwwwwwwwwwwwww").Value 1
    app.DatePicker("due").AtRect(24, 64, 90, 24).DateFormat("dddd d mmmm yyyy") _
        .PickDate DateSerial(2026, 9, 30)
    app.ComboBox("find").AtRect(24, 104, 110, 22).Items "A"
    app.Render
    app.ComboBox("find").Focus
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    RdxKeyChar "W"
    For Each faceName In Array("pick", "due", "find")
        With host.Shapes("rdm_wid79_" & faceName).TextFrame2.TextRange
            transcript = transcript & "|" & faceName & "=" & CStr(.BoundLeft + .BoundWidth _
                <= host.Shapes("rdm_wid79_" & faceName & "__caret").Left + 0.5)
        End With
    Next faceName
    RdxReleaseKeys
    ReDimUI.AutoPump True
    TestFaceClearsCaret = Mid$(transcript, 2)
End Function
