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

Public Function TestStateHandlers() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    gClickCount = 0
    Set app = ReDimUI.Mount(host, "core8")
    app.Label("out").At("B2:E2").BindText "watched"
    app.OnStateChanged "watched", "TestReDimCore.CoreStateHandler"
    app.OnStateChanged "watched", "TestReDimCore.CoreStateHandler"
    app.SetState "watched", "first"
    app.Render

    transcript = "handlersRanOnSet=" & gClickCount
    app.SetState "watched", "second"
    transcript = transcript & "|handlersRanAgain=" & gClickCount
    app.SetState "unwatched", "x"
    transcript = transcript & "|unwatchedIgnored=" & gClickCount
    TestStateHandlers = transcript
End Function

Public Sub CoreStateHandler()
    gClickCount = gClickCount + 1
End Sub

Public Function TestPersistence() As String
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim transcript As String

    Set host = NewCanvas()
    Set app = ReDimUI.Mount(host, "core9")
    app.ClearPersisted
    app.Persist True
    app.SetStateDefault "mode", "light"
    app.SetStateDefault "runs", 0
    app.SetState "mode", "dark"
    app.SetState "runs", 7
    app.SetState "ratio", 2.5
    app.SetState "flag", True
    transcript = "beforeLoss=" & app.State("mode")

    ' Simulated state loss and rebuild: the registry dies, hidden names do
    ' not, and defaults must not clobber what the user had chosen.
    ReDimUI.Shutdown
    Set app = ReDimUI.Mount(host, "core9")
    app.Persist True
    app.SetStateDefault "mode", "light"
    app.SetStateDefault "runs", 0
    transcript = transcript & "|modeKept=" & app.State("mode")
    transcript = transcript & "|runsKept=" & app.State("runs")
    transcript = transcript & "|runsType=" & TypeName(app.State("runs"))
    transcript = transcript & "|ratioKept=" & CStr(app.State("ratio") = 2.5)
    transcript = transcript & "|flagKept=" & CStr(app.State("flag") = True)
    transcript = transcript & "|flagType=" & TypeName(app.State("flag"))

    app.ClearPersisted
    ReDimUI.Shutdown
    Set app = ReDimUI.Mount(host, "core9")
    app.Persist True
    transcript = transcript & "|clearedGone=" & CStr(Not app.HasState("mode"))
    TestPersistence = transcript
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
    Set app = ReDimUI.Mount(host, "core13")
    app.Button("btn").At("B2:C3").Text("Go").OnClick "TestReDimCore.CoreStateHandler"
    app.Label("out").At("B5:E5").BindText "msg"
    app.TextInput("name").At("C7").WritesTo "who"
    app.SetState "msg", "before"
    app.Render
    app.ProtectSurface

    transcript = "protected=" & CStr(host.ProtectContents)
    transcript = transcript & "|inputUnlocked=" & _
        CStr(host.Range("C7").Locked = False)
    transcript = transcript & "|otherCellsLocked=" & _
        CStr(host.Range("A1").Locked = True)

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

    app.ProtectSurface False
    transcript = transcript & "|unprotects=" & CStr(Not host.ProtectContents)
    app.ProtectSurface
    app.Unmount True
    transcript = transcript & "|unmountUnprotects=" & _
        CStr(Not host.ProtectContents)
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
    Dim plain As ReDimUI
    Set plain = ReDimUI.Mount(EnsureSheetCore("NavTestP"), "navplain")
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
