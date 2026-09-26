Attribute VB_Name = "TestReDimCore"
Option Explicit

' Live scenarios for the ReDim core: mount, render, diffing, dispatch, state.
' Each public function is one complete scenario returning a transcript string,
' because module globals do not survive across harness round trips.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private gClickCount As Long
Private gLastSenderId As String
Private gLastSenderApp As String
Private gNavLog As String
Private gListenerSaw As String

Private Function NewCanvas() As Worksheet
    Set NewCanvas = ActiveWorkbook.Worksheets.Add
End Function

Private Function CountAppShapes( _
    ByVal host As Worksheet, _
    ByVal appId As String _
) As Long
    Dim target As Shape
    Dim total As Long

    For Each target In host.Shapes
        If Left$(target.Name, Len("rdm_" & appId & "_")) = "rdm_" & appId & "_" Then
            total = total + 1
        End If
    Next target
    CountAppShapes = total
End Function

Public Sub CoreClickHandler()
    gClickCount = gClickCount + 1
    gLastSenderId = ReDimUI.SenderId
    gLastSenderApp = ReDimUI.SenderApp.AppId
End Sub

Public Function TestMountAndRender() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim buttonShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core1")
    app.Button("run").At("B2:C3").Text("Run").Primary.OnClick "TestReDimCore.CoreClickHandler"
    app.Label("status").At("B5:E5").Text("Ready")
    app.Card("panel").AtRect(300, 20, 220, 120).Text("Card body")
    app.Render

    Set buttonShape = host.Shapes("rdm_core1_run")
    transcript = "shapes=" & CountAppShapes(host, "core1")
    transcript = transcript & "|buttonText=" & _
        buttonShape.TextFrame2.TextRange.Text
    transcript = transcript & "|fillIsPrimary=" & _
        CStr(buttonShape.Fill.ForeColor.RGB = app.Theme.PrimaryColor)
    transcript = transcript & "|onAction=" & buttonShape.OnAction
    transcript = transcript & "|labelText=" & _
        host.Shapes("rdm_core1_status").TextFrame2.TextRange.Text
    transcript = transcript & "|geometry=" & _
        CStr(Abs(buttonShape.Left - host.Range("B2").Left) < 0.01 And _
             Abs(buttonShape.Width - host.Range("B2:C3").Width) < 0.01)
    TestMountAndRender = transcript
End Function

Public Function TestIdempotentRemount() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim firstCount As Long
    Dim secondCount As Long
    Dim thirdCount As Long

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core2")
    app.Button("go").At("B2:C3").Text("Go").Primary
    app.Label("out").At("B5:E5").Text("First")
    app.Render
    firstCount = CountAppShapes(host, "core2")

    ' Same-session remount: same registry, same components.
    Set app = ReDimUI.Mount(host, "core2")
    app.Button("go").Text("Go again")
    app.Render
    secondCount = CountAppShapes(host, "core2")

    ' Simulated state loss: registry gone, shapes remain, setup reruns.
    ReDimUI.Shutdown
    Set app = ReDimUI.Mount(host, "core2")
    app.Button("go").At("B2:C3").Text("Go rebuilt").Primary
    app.Label("out").At("B5:E5").Text("Rebuilt")
    app.Render
    thirdCount = CountAppShapes(host, "core2")

    TestIdempotentRemount = "first=" & firstCount & "|second=" & secondCount & _
        "|third=" & thirdCount & "|rebuiltText=" & _
        host.Shapes("rdm_core2_go").TextFrame2.TextRange.Text
End Function

Public Function TestDispatchAndSender() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    gLastSenderId = vbNullString
    gLastSenderApp = vbNullString
    Set app = ReDimUI.Mount(host, "core3")
    app.Button("btnGo").At("B2:C3").Text("Go").OnClick "TestReDimCore.CoreClickHandler"
    app.Render

    ReDimUI.DispatchShape "rdm_core3_btnGo"
    transcript = "clicks=" & gClickCount
    transcript = transcript & "|senderId=" & gLastSenderId
    transcript = transcript & "|senderApp=" & gLastSenderApp
    transcript = transcript & "|senderCleared=" & _
        CStr(ReDimUI.Sender Is Nothing)
    ReDimUI.DispatchShape "rdm_core3_missing"
    transcript = transcript & "|missingIgnored=" & CStr(gClickCount = 1)
    TestDispatchAndSender = transcript
End Function

Public Function TestClickGuards() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    Set app = ReDimUI.Mount(host, "core4")
    app.Button("btn").At("B2:C3").Text("Guarded").OnClick "TestReDimCore.CoreClickHandler"
    app.Render

    ' Debounce: the second immediate click is ignored.
    ReDimUI.DispatchShape "rdm_core4_btn"
    ReDimUI.DispatchShape "rdm_core4_btn"
    transcript = "afterDoubleClick=" & gClickCount
    Sleep 200
    ReDimUI.DispatchShape "rdm_core4_btn"
    transcript = transcript & "|afterDebounceWait=" & gClickCount

    ' Disabled: guard blocks and the fill switches to the muted color.
    app.Button("btn").Enabled False
    Sleep 200
    ReDimUI.DispatchShape "rdm_core4_btn"
    transcript = transcript & "|afterDisabledClick=" & gClickCount
    transcript = transcript & "|disabledFillMuted=" & _
        CStr(host.Shapes("rdm_core4_btn").Fill.ForeColor.RGB = _
            app.Theme.MutedColor)
    app.Button("btn").Enabled True
    Sleep 200
    ReDimUI.DispatchShape "rdm_core4_btn"
    transcript = transcript & "|afterReEnabled=" & gClickCount
    TestClickGuards = transcript
End Function

Public Function TestStateBindings() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    Set app = ReDimUI.Mount(host, "core5")
    app.Label("msg").At("B2:E2").BindText "statusMsg", "Status: {0}"
    app.Button("btn").At("B4:C5").Text("Run").BindEnabled("canRun") _
        .OnClick "TestReDimCore.CoreClickHandler"
    app.Label("hint").At("B7:E7").Text("Hint").BindVisible "showHint"
    app.SetState "statusMsg", "starting"
    app.SetState "canRun", False
    app.SetState "showHint", True
    app.Render

    transcript = "boundText=" & _
        host.Shapes("rdm_core5_msg").TextFrame2.TextRange.Text
    ReDimUI.DispatchShape "rdm_core5_btn"
    transcript = transcript & "|disabledClicks=" & gClickCount

    app.SetState "statusMsg", "ready"
    app.SetState "canRun", True
    app.SetState "showHint", False
    transcript = transcript & "|updatedText=" & _
        host.Shapes("rdm_core5_msg").TextFrame2.TextRange.Text
    transcript = transcript & "|hintHidden=" & _
        CStr(host.Shapes("rdm_core5_hint").Visible = msoFalse)
    ReDimUI.DispatchShape "rdm_core5_btn"
    transcript = transcript & "|enabledClicks=" & gClickCount
    transcript = transcript & "|stateReadback=" & app.State("statusMsg")

    ' Inverse binding: enabled while the busy flag is False.
    app.Button("inv").At("B9:C10").Text("Inverse").BindEnabled "busy", True
    app.SetState "busy", False
    app.Render
    transcript = transcript & "|invertedIdleEnabled=" & _
        CStr(app.Button("inv").IsEnabled)
    app.SetState "busy", True
    transcript = transcript & "|invertedBusyDisabled=" & _
        CStr(Not app.Button("inv").IsEnabled)
    TestStateBindings = transcript
End Function

Public Function TestBatchAndTheme() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core6")
    app.Button("btn").At("B2:C3").Text("Styled").Primary
    app.Label("lbl").At("B5:E5").BindText "message"
    app.SetState "message", "before"
    app.Render

    app.BeginUpdate
    app.SetState "message", "step1"
    app.SetState "message", "step2"
    app.SetState "message", "batched"
    app.EndUpdate
    transcript = "batchedText=" & _
        host.Shapes("rdm_core6_lbl").TextFrame2.TextRange.Text

    app.SetTheme ReDimUI.ThemeDark
    transcript = transcript & "|darkFill=" & _
        CStr(host.Shapes("rdm_core6_btn").Fill.ForeColor.RGB = _
            app.Theme.PrimaryColor)
    transcript = transcript & "|darkIsDark=" & _
        CStr(app.Theme.PrimaryColor <> ReDimUI.ThemeLight.PrimaryColor)
    TestBatchAndTheme = transcript
End Function

' A theme swap must repaint everything already on screen, including the
' painted canvas background and once-styled fills like the progress
' track, not just the components whose colors happen to be diffed.
Public Function TestThemeSwapRepaint() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "corethm")
    app.PrepareCanvas
    app.Card("panel").AtRect(20, 20, 200, 90).Text("Panel")
    app.ProgressBar("bar").AtRect(30, 130, 160, 10).Value 40
    app.Render
    transcript = "lightCanvas=" & _
        CStr(host.Cells(1, 1).Interior.Color = app.Theme.CanvasColor)

    app.SetTheme ReDimUI.ThemeDark
    transcript = transcript & "|darkCanvas=" & _
        CStr(host.Cells(1, 1).Interior.Color = app.Theme.CanvasColor)
    transcript = transcript & "|darkCard=" & _
        CStr(host.Shapes("rdm_corethm_panel").Fill.ForeColor.RGB = _
            app.Theme.SurfaceColor)
    transcript = transcript & "|darkTrack=" & _
        CStr(host.Shapes("rdm_corethm_bar").Fill.ForeColor.RGB = _
            app.Theme.MutedColor)
    transcript = transcript & "|swapChanged=" & _
        CStr(app.Theme.CanvasColor <> ReDimUI.ThemeLight.CanvasColor)
    TestThemeSwapRepaint = transcript
End Function

Public Function TestStateHandlers() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    Set app = ReDimUI.Mount(host, "core8")
    app.Label("out").At("B2:E2").BindText "watched"
    ' The same listener twice, as a build that runs twice registers it,
    ' counts once; a second listener watches two keys from one call.
    app.OnStateChanged "watched", "TestReDimCore.CoreStateHandler"
    app.OnStateChanged "watched", "TestReDimCore.CoreStateHandler"
    app.OnStateChanged Array("watched", "other"), "TestReDimCore.CoreStateHandlerTens"
    app.SetState "watched", "first"
    app.Render

    transcript = "handlersRanOnSet=" & gClickCount
    app.SetState "watched", "second"
    transcript = transcript & "|handlersRanAgain=" & gClickCount
    app.SetState "unwatched", "x"
    transcript = transcript & "|unwatchedIgnored=" & gClickCount
    app.SetState "other", "y"
    transcript = transcript & "|secondKeyHeard=" & gClickCount
    TestStateHandlers = transcript
End Function

Public Sub CoreStateHandler()
    gClickCount = gClickCount + 1
End Sub

Public Sub CoreStateHandlerTens()
    gClickCount = gClickCount + 10
End Sub

' State is session-scoped by design: a rebuild starts empty, and
' SetStateDefault only seeds keys that have no value yet.
Public Function TestStateSessionScope() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core9")
    app.SetStateDefault "mode", "light"
    app.SetState "mode", "dark"
    app.SetStateDefault "mode", "light"
    transcript = "defaultNoClobber=" & app.State("mode")

    ReDimUI.Shutdown
    Set app = ReDimUI.Mount(host, "core9")
    transcript = transcript & "|freshStoreEmpty=" & _
        CStr(Not app.HasState("mode"))
    app.SetStateDefault "mode", "light"
    transcript = transcript & "|defaultSeeds=" & app.State("mode")
    TestStateSessionScope = transcript
End Function

Public Function TestRelativeLayout() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim titleShape As Shape
    Dim underShape As Shape
    Dim asideShape As Shape
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core10")
    app.Label("hdr").AtRect(30, 20, 200, 24).Text("Header")
    app.Button("under").Below("hdr", 10).Sized(120, 30).Text("Under")
    app.Button("aside").RightOf("under", 12).Sized(90, 30).Text("Aside")
    app.Render

    Set titleShape = host.Shapes("rdm_core10_hdr")
    Set underShape = host.Shapes("rdm_core10_under")
    Set asideShape = host.Shapes("rdm_core10_aside")
    transcript = "underLeftAligned=" & _
        CStr(Abs(underShape.Left - titleShape.Left) < 0.01)
    transcript = transcript & "|underBelow=" & _
        CStr(Abs(underShape.Top - (titleShape.Top + titleShape.Height + 10)) < 0.01)
    transcript = transcript & "|asideTopAligned=" & _
        CStr(Abs(asideShape.Top - underShape.Top) < 0.01)
    transcript = transcript & "|asideRight=" & _
        CStr(Abs(asideShape.Left - (underShape.Left + underShape.Width + 12)) < 0.01)

    On Error Resume Next
    app.Button("bad").Below "missing"
    app.Render
    transcript = transcript & "|badRefErr=" & CStr(Err.Number <> 0)
    On Error GoTo 0
    app.Button("bad").Remove
    TestRelativeLayout = transcript
End Function

Public Function TestOrphanPruning() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core11")
    app.Button("oldname").At("B2:C3").Text("Old")
    app.Label("keeper").At("B5:D5").Text("Keeper")
    app.Render
    transcript = "beforeCount=" & CountAppShapes(host, "core11")

    ' Rebuild after state loss with a renamed component: the stale shape
    ' must be swept, the surviving one adopted.
    ReDimUI.Shutdown
    Set app = ReDimUI.Mount(host, "core11")
    app.Button("newname").At("B2:C3").Text("New")
    app.Label("keeper").At("B5:D5").Text("Keeper")
    app.Render
    transcript = transcript & "|afterCount=" & CountAppShapes(host, "core11")
    transcript = transcript & "|oldGone=" & _
        CStr(Not ShapeExistsCore(host, "rdm_core11_oldname"))
    transcript = transcript & "|newExists=" & _
        CStr(ShapeExistsCore(host, "rdm_core11_newname"))
    TestOrphanPruning = transcript
End Function

Public Function TestHotKeyLifecycle() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    Set app = ReDimUI.Mount(host, "core12")
    ' CoreStateHandler is context-free; CoreClickHandler reads the sender,
    ' which only exists during dispatch.
    app.HotKey "^+{F12}", "TestReDimCore.CoreStateHandler"
    Application.Run "TestReDimCore.CoreStateHandler"
    transcript = "procCallable=" & CStr(gClickCount = 1)
    app.Unmount True
    transcript = transcript & "|unmountClean=True"
    transcript = transcript & "|version=" & ReDimUI.Version
    TestHotKeyLifecycle = transcript
End Function

Public Function TestProtectSurface() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    ReDimUI.AutoPump False
    Set app = ReDimUI.Mount(host, "core13")
    app.Button("btn").At("B2:C3").Text("Go").OnClick "TestReDimCore.CoreStateHandler"
    app.Label("out").At("B5:E5").BindText "msg"
    app.TextInput("name").At("C7").WritesTo "who"
    app.SelectBox("pick").AtRect(300, 20, 120, 22).Items "A", "B", "C", "D", "E", "F"
    app.DatePicker("due").AtRect 450, 20, 140, 22
    app.SetState "msg", "before"
    app.Render
    app.ProtectSurface

    transcript = "protected=" & CStr(host.ProtectContents)
    transcript = transcript & "|inputUnlocked=" & _
        CStr(host.Range("C7").Locked = False)
    transcript = transcript & "|otherCellsLocked=" & _
        CStr(host.Range("A1").Locked = True)
    ' Locked canvas cells are unselectable by default; the unlocked input
    ' cell remains selectable so cell-backed typing still works.
    transcript = transcript & "|selectionLocked=" & _
        CStr(host.EnableSelection = xlUnlockedCells)

    ' Framework writes keep working under UserInterfaceOnly protection.
    app.SetState "msg", "updated under protection"
    transcript = transcript & "|renderWorks=" & _
        CStr(host.Shapes("rdm_core13_out").TextFrame2.TextRange.Text = _
            "updated under protection")
    ReDimUI.DispatchShape "rdm_core13_btn"
    transcript = transcript & "|dispatchWorks=" & CStr(gClickCount = 1)
    Dim toastValue As ReDimUI
    Set toastValue = app.Toast("under protection", 60000)
    transcript = transcript & "|toastCreates=" & _
        CStr(Not toastValue Is Nothing)
    ' A drop list and a calendar open under protection: their rows and
    ' weeks copy with the protection lifted for a moment, and it comes
    ' back as it was, selection rule included.
    ReDimUI.DispatchShape "rdm_core13_pick"
    transcript = transcript & "|protectedListOpens=" & CStr( _
        ShapeExistsCore(host, "rdm_core13_pick__opt6") And host.ProtectDrawingObjects _
        And host.ProtectContents And host.ProtectionMode _
        And host.EnableSelection = xlUnlockedCells)
    ReDimUI.DispatchShape "rdm_core13_due"
    transcript = transcript & "|protectedCalendarOpens=" & CStr( _
        ShapeExistsCore(host, "rdm_core13_due__cr6") And host.ProtectDrawingObjects _
        And host.ProtectContents And host.ProtectionMode _
        And host.EnableSelection = xlUnlockedCells)
    ReDimUI.DispatchShape "rdm_core13_due"

    app.ProtectSurface False
    transcript = transcript & "|unprotects=" & CStr(Not host.ProtectContents)
    transcript = transcript & "|selectionRestored=" & _
        CStr(host.EnableSelection = xlNoRestrictions)
    app.ProtectSurface True, True
    transcript = transcript & "|optOutSelectable=" & _
        CStr(host.EnableSelection = xlNoRestrictions)
    app.ProtectSurface False
    app.ProtectSurface
    app.Unmount True
    transcript = transcript & "|unmountUnprotects=" & _
        CStr(Not host.ProtectContents)
    ReDimUI.AutoPump True
    TestProtectSurface = transcript
End Function

Public Sub NavShowA()
    gNavLog = gNavLog & "+A"
End Sub

Public Sub NavHideA()
    gNavLog = gNavLog & "-A"
End Sub

Public Sub NavShowB()
    gNavLog = gNavLog & "+B"
End Sub

Private Function EnsureSheetCore(ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In ActiveWorkbook.Worksheets
        If candidate.Name = sheetName Then
            Set EnsureSheetCore = candidate
            Exit Function
        End If
    Next candidate
    Set EnsureSheetCore = ActiveWorkbook.Worksheets.Add
    EnsureSheetCore.Name = sheetName
End Function

Public Function TestNavigation() As String
    Dim appA As ReDimUI
    Dim appB As ReDimUI
    Dim appC As ReDimUI
    Dim sheetA As Worksheet
    Dim sheetB As Worksheet
    Dim sheetC As Worksheet
    Dim transcript As String

    gNavLog = vbNullString
    Set sheetA = EnsureSheetCore("NavTestA")
    Set sheetB = EnsureSheetCore("NavTestB")
    Set sheetC = EnsureSheetCore("NavTestC")
    sheetA.Visible = xlSheetVisible
    sheetB.Visible = xlSheetVisible
    sheetC.Visible = xlSheetVisible

    Set appA = ReDimUI.Mount(sheetA, "nava")
    appA.AsWindow.OnShow "TestReDimCore.NavShowA"
    appA.OnHide "TestReDimCore.NavHideA"
    appA.Button("gob").At("B2:C3").Text("Go B").NavigatesTo "navb"
    appA.Render
    Set appB = ReDimUI.Mount(sheetB, "navb")
    appB.AsWindow.OnShow "TestReDimCore.NavShowB"
    Set appC = ReDimUI.Mount(sheetC, "navc")
    appC.AsWindow
    appA.WindowTitle "Alpha"
    appA.NavBar

    ReDimUI.Navigate "nava"
    transcript = "activeA=" & CStr(ReDimUI.ActiveWindowId = "nava")
    transcript = transcript & "|aVisible=" & _
        CStr(sheetA.Visible = xlSheetVisible)
    transcript = transcript & "|bHidden=" & _
        CStr(sheetB.Visible = xlSheetVeryHidden)
    transcript = transcript & "|cHidden=" & _
        CStr(sheetC.Visible = xlSheetVeryHidden)
    transcript = transcript & "|navBarTabs=" & _
        CStr(ShapeExistsCore(sheetA, "rdm_nava_nvb_nava") And _
             ShapeExistsCore(sheetA, "rdm_nava_nvb_navb") And _
             ShapeExistsCore(sheetA, "rdm_nava_nvb_navc"))
    transcript = transcript & "|activeTabPrimary=" & _
        CStr(sheetA.Shapes("rdm_nava_nvb_nava").Fill.ForeColor.RGB = _
            appA.Theme.PrimaryColor)
    transcript = transcript & "|tabTitleText=" & _
        sheetA.Shapes("rdm_nava_nvb_nava").TextFrame2.TextRange.Text

    ' A button link navigates and fires lifecycle hooks in order.
    ReDimUI.DispatchShape "rdm_nava_gob"
    transcript = transcript & "|activeB=" & _
        CStr(ReDimUI.ActiveWindowId = "navb")
    transcript = transcript & "|aNowHidden=" & _
        CStr(sheetA.Visible = xlSheetVeryHidden)
    transcript = transcript & "|log=" & gNavLog
    transcript = transcript & "|tabSwapped=" & _
        CStr(sheetA.Shapes("rdm_nava_nvb_navb").Fill.ForeColor.RGB = _
            appA.Theme.PrimaryColor)

    ReDimUI.Navigate "navc"
    transcript = transcript & "|activeC=" & _
        CStr(ReDimUI.ActiveWindowId = "navc")
    transcript = transcript & "|backToB=" & CStr(ReDimUI.NavigateBack())
    transcript = transcript & "|activeAfterBack=" & ReDimUI.ActiveWindowId
    transcript = transcript & "|backToA=" & CStr(ReDimUI.NavigateBack())
    transcript = transcript & "|backEmpty=" & CStr(Not ReDimUI.NavigateBack())

    ' Non-window apps refuse navigation with a clear error.
    ReDimUI.Mount EnsureSheetCore("NavTestP"), "navplain"
    On Error Resume Next
    Err.Clear
    ReDimUI.Navigate "navplain"
    transcript = transcript & "|plainRefused=" & CStr(Err.Number <> 0)
    On Error GoTo 0

    ' Unmounting a window forgets it, and the next navigation prunes its
    ' tab from every bar.
    appC.Unmount True
    On Error Resume Next
    Err.Clear
    ReDimUI.Navigate "navc"
    transcript = transcript & "|goneRefused=" & CStr(Err.Number <> 0)
    On Error GoTo 0
    ReDimUI.Navigate "navb"
    transcript = transcript & "|staleTabPruned=" & _
        CStr(Not ShapeExistsCore(sheetA, "rdm_nava_nvb_navc"))
    transcript = transcript & "|liveTabsRemain=" & _
        CStr(ShapeExistsCore(sheetA, "rdm_nava_nvb_nava") And _
             ShapeExistsCore(sheetA, "rdm_nava_nvb_navb"))

    ' Shutdown forgets the windows with their apps: a rebuild in the same
    ' session gets a bar with no tabs for unmounted windows, no active
    ' window, and an empty back stack.
    ReDimUI.Shutdown
    transcript = transcript & "|shutdownClearsActive=" & _
        CStr(LenB(ReDimUI.ActiveWindowId) = 0)
    On Error Resume Next
    Err.Clear
    Set appA = ReDimUI.Mount(sheetA, "nava")
    appA.AsWindow.NavBar
    transcript = transcript & "|rebuildAfterShutdown=" & CStr(Err.Number = 0)
    Err.Clear
    transcript = transcript & "|backStackCleared=" & _
        CStr(Not ReDimUI.NavigateBack() And Err.Number = 0)
    On Error GoTo 0
    ReDimUI.Shutdown
    TestNavigation = transcript
End Function

Private Function ShapeExistsCore( _
    ByVal host As Worksheet, _
    ByVal shapeName As String _
) As Boolean
    Dim probe As Shape

    On Error Resume Next
    Set probe = host.Shapes(shapeName)
    On Error GoTo 0
    ShapeExistsCore = Not probe Is Nothing
End Function

Public Function TestUnmount() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim beforeCount As Long

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core7")
    app.Button("a").At("B2:C3").Text("A")
    app.Label("b").At("B5:E5").Text("B")
    app.Render
    beforeCount = CountAppShapes(host, "core7")
    app.Unmount True
    TestUnmount = "before=" & beforeCount & _
        "|after=" & CountAppShapes(host, "core7") & _
        "|forgotten=" & CStr(Not ReDimUI.HasApp("core7"))
End Function

' A control follows the key it writes. A key with no value takes the
' control's on the first render, silently; a key with a value shows on
' the control after SetState, firing no OnChange; a select maps a value
' or text back to its item; a focused field keeps what is being typed;
' and a Render shows the value in play over the one a rebuild declares.
Public Function TestWritesToFollows() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    gClickCount = 0
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core10")
    app.OnStateChanged "darkMode", "TestReDimCore.CoreStateHandler"
    app.Toggle("dark").AtRect(24, 24, 44, 22).Checked(True).WritesTo("darkMode") _
        .OnChange "TestReDimCore.CoreStateHandlerTens"
    app.SelectBox("size").AtRect(24, 60, 140, 22).WritesTo "size"
    app.SelectBox("size").AddItem "Small"
    app.SelectBox("size").AddItem "Medium", , 20
    app.Stepper("qty").AtRect(24, 100, 120, 24).SliderRange(0, 50, 1).WritesTo "qty"
    app.TextInput("name").AtRect(24, 140, 150, 22).WritesTo "name"
    app.Render
    transcript = "seeded=" & CStr(app.State("darkMode") = True And app.State("qty") = 0 _
        And Not app.HasState("name") And Not app.HasState("size")) & "/" & gClickCount
    app.SetState "darkMode", False
    transcript = transcript & "|follows=" & CStr(Not app.Toggle("dark").IsChecked) & "/" & gClickCount
    app.SetState "size", 20
    transcript = transcript & "|valueMaps=" & app.SelectBox("size").CurrentValue
    app.SetState "size", "Small"
    transcript = transcript & "|textMaps=" & app.SelectBox("size").CurrentValue
    app.SetState "qty", 999
    transcript = transcript & "|clamps=" & app.Stepper("qty").CurrentValue
    app.SetState "name", "Ada"
    transcript = transcript & "|fieldFollows=" & app.TextInput("name").InputValue
    app.TextInput("name").Focus
    RdxKeyChar "x"
    app.SetState "name", "Bob"
    transcript = transcript & "|typingKept=" & app.TextInput("name").InputValue
    RdxReleaseKeys
    ReDimUI.ClearKeyboardFocus
    app.SetState "darkMode", True
    app.Toggle("dark").Checked False
    app.Render
    transcript = transcript & "|renderKeepsState=" & CStr(app.Toggle("dark").IsChecked)
    app.Unmount True
    TestWritesToFollows = transcript
End Function

' WritesTo at its edges: a label bound to a seeded key shows the seed,
' keys match in any case, a pick among items sharing a value stays on
' the item picked, Empty and Null clear a pick without breaking later
' draws, a same-key BindValue keeps the clamp, a date comes from a
' numeric string, a setter is overruled by the next write of its key,
' a listener reads the value the control now shows, and a Required
' message goes once state fills the control.
Public Function TestWritesToEdges() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core11")
    app.Label("volText").AtRect(200, 24, 120, 18).BindText "vol", "Volume {0}"
    app.RadioGroup("size").AtRect(24, 24, 140, 70).WritesTo "size"
    app.RadioGroup("size").AddItem "Small", , 1
    app.RadioGroup("size").AddItem "Tiny", , 1
    app.RadioGroup("size").AddItem "None", , 0
    app.Stepper("vol").AtRect(24, 110, 120, 24).SliderRange(0, 100, 1).Value(40) _
        .WritesTo("Vol").BindValue "Vol"
    app.DatePicker("due").AtRect(24, 150, 150, 24).WritesTo "due"
    app.Toggle("dark").AtRect(24, 190, 44, 22).WritesTo "dark"
    app.SelectBox("kind").AtRect(200, 60, 140, 22).Items("A", "B").Required.WritesTo "kind"
    app.CheckList("opts").AtRect(360, 24, 150, 80).Items("Alpha", "Bravo", "Charlie") _
        .CheckedFrom(Array("Alpha", "Charlie")).WritesTo "opts"
    app.TransferList("crew").AtRect(200, 240, 360, 140).Items("Ann", "Ben", "Cy") _
        .ChosenFrom(Array("Dee")).WritesTo "crew"
    app.OnStateChanged "dark", "TestReDimCore.CoreReadToggle"
    app.Render
    transcript = "seedShown=" & app.Label("volText").CurrentText
    transcript = transcript & "|listsSeed=" & app.State("opts") & "/" & app.State("crew")
    transcript = transcript & "|positions=" & app.TransferList("crew").ItemPosition("ben") & _
        "/" & app.TransferList("crew").ChosenPosition("DEE") & "/" & _
        app.TransferList("crew").ChosenPosition("Ann")
    ReDimUI.DispatchShape "rdm_core11_size__t2"
    transcript = transcript & "|sharedValueKept=" & app.RadioGroup("size").CurrentValue & _
        "/" & app.State("size")
    app.SetState "size", 0
    transcript = transcript & "|zeroNames=" & app.RadioGroup("size").CurrentValue
    app.SetState "size", Empty
    transcript = transcript & "|emptyClears=" & app.RadioGroup("size").CurrentValue
    app.SetState "size", 0
    app.SetState "size", Null
    transcript = transcript & "|nullClears=" & app.RadioGroup("size").CurrentValue
    app.SetState "vol", 30
    transcript = transcript & "|drawsAfterNull=" & app.Stepper("vol").CurrentValue & "/" & _
        app.Label("volText").CurrentText
    app.SetState "vol", 150
    transcript = transcript & "|clampKept=" & app.Stepper("vol").CurrentValue
    app.SetState "due", "46287"
    transcript = transcript & "|dateFromText=" & CStr(app.DatePicker("due").PickedDate = _
        CDate(46287#))
    app.SetState "due", 1E+20
    transcript = transcript & "|hugeIgnored=" & CStr(app.DatePicker("due").PickedDate = _
        CDate(46287#))
    app.SetState "dark", True
    transcript = transcript & "|listenerSees=" & gListenerSaw
    app.Toggle("dark").Checked False
    app.SetState "dark", True
    transcript = transcript & "|writeOverrules=" & CStr(app.Toggle("dark").IsChecked)
    app.ValidateAll
    RdxReleaseKeys
    ReDimUI.ClearKeyboardFocus
    app.SetState "kind", "B"
    transcript = transcript & "|requiredClears=" & _
        CStr(Not ShapeExistsCore(host, "rdm_core11_kind__me"))
    app.Unmount True
    TestWritesToEdges = transcript
End Function

' The rails around focus, dialogs, removal, and a workbook close: a
' focus handoff commits and still frees the keys at the next blur, a
' capture with nothing focused lets the keys go at once, a sheet behind
' another captures no keys until it comes to the front, a Confirm
' opened from a dialog's OK stays open and focus returns past both, a
' focused field that hides commits its edits, removing an anchor leaves
' its dependents in place, At replaces Below, an id cannot end in an
' underscore, word keys treat Hangul as letters, and a held close keeps
' the apps.
Public Function TestFocusRails() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim other As Worksheet
    Dim transcript As String
    Dim depTop As Double
    Dim eventsWereOn As Boolean

    ReDimUI.AutoPump False
    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core12")
    app.TextInput("a").AtRect(24, 24, 150, 22).WritesTo "k"
    app.TextInput("b").AtRect 24, 60, 150, 22
    app.TextInput("c").AtRect(24, 96, 150, 22).WritesTo "cText"
    app.TextInput("w").AtRect 24, 132, 150, 22
    app.Label("anchor").AtRect(240, 24, 120, 18).Text "Anchor"
    app.Label("dep").Below("anchor", 6).Sized(120, 18).Text "Dependent"
    app.Label("rel").Below("anchor", 40).Sized(120, 18).Text "Relative"
    app.Button("go").AtRect(240, 160, 90, 28).Text("Go").OnClick "TestReDimCore.CoreClickHandler"
    app.Render

    app.TextInput("a").Focus
    transcript = "captured=" & CStr(RdxKeysCaptured())
    RdxKeyChar "x"
    app.TextInput("b").Focus
    transcript = transcript & "|handoff=" & app.State("k") & "/" & ReDimUI.FocusedComponentId
    ReDimUI.ClearKeyboardFocus
    transcript = transcript & "|keysFreed=" & CStr(Not RdxKeysCaptured())

    RdxBindKeys
    RdxKeyChar "q"
    transcript = transcript & "|strayFreed=" & CStr(Not RdxKeysCaptured())

    ' The harness may hold EnableEvents off; the sheet's activation needs
    ' them on to reach ReDim.
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = True
    Set other = NewCanvas()
    app.TextInput("b").Focus
    transcript = transcript & "|behindFree=" & CStr(Not RdxKeysCaptured())
    host.Activate
    transcript = transcript & "|frontTakes=" & CStr(RdxKeysCaptured())
    Application.EnableEvents = eventsWereOn

    app.Confirm "First", "One", "TestReDimCore.CoreConfirmAgain"
    Sleep 250
    ReDimUI.DispatchShape "rdm_core12_mdl_ok"
    transcript = transcript & "|secondOpen=" & _
        CStr(host.Shapes("rdm_core12_mdl_card").Visible = msoTrue) & "/" & _
        CStr(InStr(host.Shapes("rdm_core12_mdl_card").TextFrame2.TextRange.Text, "Second") > 0)
    Sleep 250
    ReDimUI.DispatchShape "rdm_core12_mdl_ok"
    transcript = transcript & "|focusBack=" & ReDimUI.FocusedComponentId

    app.TextInput("c").Focus
    RdxKeyChar "h"
    RdxKeyChar "i"
    app.TextInput("c").Visible False
    transcript = transcript & "|hiddenCommits=" & app.State("cText") & "/" & _
        CStr(LenB(ReDimUI.FocusedComponentId) = 0)

    depTop = host.Shapes("rdm_core12_dep").Top
    app.Label("anchor").Remove
    app.Label("dep").Text "Still here"
    transcript = transcript & "|depStays=" & CStr(host.Shapes("rdm_core12_dep").Top = depTop)
    app.Label("rel").At "D12"
    transcript = transcript & "|atReplaces=" & CStr(host.Shapes("rdm_core12_rel").Top = _
        host.Range("D12").Top)

    On Error Resume Next
    app.Label "bad_"
    transcript = transcript & "|idEnds=" & Err.Description
    Err.Clear
    On Error GoTo 0

    With app.TextInput("w")
        .InputValue = ChrW(&H65E5) & ChrW(&H672C) & " " & ChrW(&HD55C) & ChrW(&HAE00)
    End With
    app.TextInput("w").Focus
    RdxKeyChar "{END}"
    RdxKeyChar "{WORDBS}"
    transcript = transcript & "|hangulWord=" & CStr(app.TextInput("w").InputValue = _
        ChrW(&H65E5) & ChrW(&H672C) & " ")
    ReDimUI.ClearKeyboardFocus

    gClickCount = 0
    ReDimUI.HoldForClose
    transcript = transcript & "|heldKeepsApp=" & CStr(ReDimUI.HasApp("core12"))
    Sleep 250
    ReDimUI.DispatchShape "rdm_core12_go"
    transcript = transcript & "|clickResumes=" & gClickCount
    app.Unmount True
    Application.DisplayAlerts = False
    other.Delete
    Application.DisplayAlerts = True
    TestFocusRails = transcript
End Function

Public Sub CoreConfirmAgain()
    ReDimUI.App("core12").Confirm "Second", "Two"
End Sub

Public Sub CoreReadToggle()
    gListenerSaw = CStr(ReDimUI.App("core11").Toggle("dark").IsChecked)
End Sub

' The theme builders set every color a theme holds, each read back by
' its token, and building on a preset leaves the next preset untouched.
Public Function TestThemeBuilders() As String
    Dim themeValue As ReDimUI
    Dim transcript As String

    Set themeValue = ReDimUI.ThemeLight.WithSurface(RGB(1, 2, 3), RGB(4, 5, 6)) _
        .WithMuted(RGB(7, 8, 9), RGB(10, 11, 12)).WithStatus(RGB(13, 14, 15), RGB(16, 17, 18)) _
        .WithBorder(RGB(19, 20, 21)).WithCanvas(RGB(22, 23, 24))
    transcript = "tokens=" & CStr(themeValue.SurfaceColor = RGB(1, 2, 3) _
        And themeValue.OnSurfaceColor = RGB(4, 5, 6) _
        And themeValue.MutedColor = RGB(7, 8, 9) _
        And themeValue.OnMutedColor = RGB(10, 11, 12) _
        And themeValue.SuccessColor = RGB(13, 14, 15) _
        And themeValue.DangerColor = RGB(16, 17, 18) _
        And themeValue.BorderColor = RGB(19, 20, 21) _
        And themeValue.CanvasColor = RGB(22, 23, 24))
    transcript = transcript & "|stockUntouched=" & CStr(ReDimUI.ThemeLight.SurfaceColor = _
        RGB(255, 255, 255))
    TestThemeBuilders = transcript
End Function

' Errors say where they came from: a component's lead with its id, a
' kind clash names both kinds, a member on the wrong role names the role
' it needs, and a relative placement names the builder that set it.
Public Function TestErrorWords() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core8")
    app.SelectBox("zone").AtRect 24, 24, 120, 22
    app.Toggle("dark").AtRect 24, 60, 44, 22
    On Error Resume Next
    app.SelectBox("zone").RestrictToItems
    transcript = "componentNamed=" & Err.Description
    Err.Clear
    app.Button "dark"
    transcript = transcript & "|kindClash=" & Err.Description
    Err.Clear
    app.Text "x"
    transcript = transcript & "|roleNamed=" & Err.Description
    Err.Clear
    app.Label("tag").RightOf "nowhere"
    app.Render
    transcript = transcript & "|relNamed=" & Err.Description
    Err.Clear
    app.Toggle("dark").Clearable
    transcript = transcript & "|fieldOnly=" & Err.Description
    Err.Clear
    On Error GoTo 0
    app.Unmount True
    TestErrorWords = transcript
End Function
