VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReDimUI"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ReDim: a stateful UI framework for Excel worksheets, built on ROneCOne.
' One role-tagged class in the ROneCOne style. The predeclared instance is the
' factory and app registry; other roles are apps, components, themes, async
' ops, and jobs. ReDimHost.bas supplies the OnAction and SetTimer entry points
' that Excel can only route to a standard module.

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef counter As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef frequency As LongLong) As Long
Private Type RdxScreenPoint
    X As Long
    Y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
    ByRef position As RdxScreenPoint) As Long
Private Declare PtrSafe Function GetDC Lib "user32" ( _
    ByVal windowHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal windowHandle As LongPtr, ByVal deviceContext As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
    ByVal deviceContext As LongPtr, ByVal capabilityIndex As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal virtualKey As Long) As Integer

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const VK_LEFT_BUTTON As Long = 1
Private Const VK_RIGHT_BUTTON As Long = 2

Private Const SHAPE_PREFIX As String = "rdm_"
Private Const PART_SEPARATOR As String = "__"
Private Const ERROR_BASE As Long = vbObjectError + 4200
Private Const ERROR_INVALID_ARGUMENT As Long = ERROR_BASE + 1
Private Const ERROR_INVALID_OPERATION As Long = ERROR_BASE + 2
Private Const ERROR_NOT_FOUND As Long = ERROR_BASE + 3
Private Const CLICK_DEBOUNCE_MS As Long = 150
Private Const UNMAPPABLE_POINTS As Double = -1E+18

Private Const ROLE_UNSET As Long = 0
Private Const ROLE_APP As Long = 1
Private Const ROLE_COMPONENT As Long = 2
Private Const ROLE_THEME As Long = 3
Private Const ROLE_ASYNC_OP As Long = 4
Private Const ROLE_JOB As Long = 5

Private Const KIND_NONE As Long = 0
Private Const KIND_BUTTON As Long = 1
Private Const KIND_LABEL As Long = 2
Private Const KIND_CARD As Long = 3
Private Const KIND_PROGRESS As Long = 4
Private Const KIND_SPINNER As Long = 5
Private Const KIND_TOGGLE As Long = 6
Private Const KIND_INPUT As Long = 10
Private Const KIND_TOAST As Long = 11
Private Const KIND_OVERLAY As Long = 12
Private Const KIND_SELECT As Long = 13
Private Const KIND_TICKBOX As Long = 14
Private Const KIND_RADIO As Long = 15
Private Const KIND_STEPPER As Long = 16
Private Const KIND_SLIDEBAR As Long = 17
Private Const KIND_COMBO As Long = 18

Private Const VARIANT_PRIMARY As Long = 0
Private Const VARIANT_SECONDARY As Long = 1
Private Const VARIANT_SUCCESS As Long = 2
Private Const VARIANT_DANGER As Long = 3
Private Const VARIANT_NEUTRAL As Long = 4

Private Const REL_NONE As Long = 0
Private Const REL_BELOW As Long = 1
Private Const REL_RIGHT_OF As Long = 2
Private Const PERSIST_PREFIX As String = "rdm_s_"
Private Const REDIM_VERSION As String = "0.9.2"

Private mRole As Long

' Factory fields (meaningful only on the predeclared instance).
Private mApps As Collection
Private mSenderApp As ReDimUI
Private mSenderComponent As ReDimUI
Private mSenderPartName As String
Private mAutoPumpOff As Boolean
Private mInPumpTick As Boolean
Private mLastTickPreciseMs As Double
Private mWorkDebtMs As Long
Private mWindows As Collection
Private mNavStack As Collection
Private mActiveWindowId As String
Private mFocusAppId As String
Private mFocusComponentId As String
Private WithEvents mExcelApp As Excel.Application
Attribute mExcelApp.VB_VarHelpID = -1

' App fields.
Private mAppId As String
Private mSheet As Worksheet
Private mComponents As Collection
Private mComponentOrder As Collection
Private mStateValues As Collection
Private mStateHandlers As Collection
Private mDirtyIds As Collection
Private mBatchDepth As Long
Private mRendered As Boolean
Private mTheme As ReDimUI
Private mAsyncOps As Collection
Private mJobs As Collection
Private mErrorHandlerProc As String
Private mToastCounter As Long
Private mFieldPrevButtonDown As Boolean
Private mToastTrayAnchor As String
Private mPersistOn As Boolean
Private mHotKeys As Collection
Private mSurfaceProtected As Boolean
Private mIsWindow As Boolean
Private mOnShowProc As String
Private mOnHideProc As String
Private mCanvasPrepared As Boolean
Private mWindowTitle As String
Private mHasNavBar As Boolean
Private mNavBarLeft As Double
Private mNavBarTop As Double
Private mNavBarTabWidth As Double
Private mNavBarTabHeight As Double

' Component fields.
Private mComponentId As String
Private mOwnerAppId As String
Private mKind As Long
Private mAnchor As String
Private mLeft As Double
Private mTop As Double
Private mWidth As Double
Private mHeight As Double
Private mHasRect As Boolean
Private mText As String
Private mBusyText As String
Private mStyleVariant As Long
Private mFillColor As Long
Private mTextColor As Long
Private mFontSize As Double
Private mFontBold As Boolean
Private mVisible As Boolean
Private mEnabled As Boolean
Private mBusy As Boolean
Private mValue As Double
Private mChecked As Boolean
Private mClickProc As String
Private mAsyncClickProc As String
Private mNavTargetId As String
Private mBindTextKey As String
Private mBindTextFormat As String
Private mBindValueKey As String
Private mBindVisibleKey As String
Private mBindVisibleInvert As Boolean
Private mBindEnabledKey As String
Private mBindEnabledInvert As Boolean
Private mWritesToKey As String
Private mLastClickMs As LongLong
Private mExpiresAtMs As LongLong
Private mSlideRemaining As Double
Private mItems As Collection
Private mSliderMin As Long
Private mSliderMax As Long
Private mSliderStep As Long
Private mCloseModalOnClick As Boolean
Private mSelectOpen As Boolean
Private mAppliedOptionCount As Long
Private mSlideTracking As Boolean
Private mSlidePrevButtonDown As Boolean
Private mSlideValueAtEngage As Double
Private mSlideSuppressUntilMs As LongLong
Private mComboMatches As Collection
Private mFieldTextAtFocus As String
Private mCaretOn As Boolean
Private mCaretAccumMs As Long
Private mRelMode As Long
Private mRelTargetId As String
Private mRelGap As Double
Private mGeometryResolving As Boolean

' Applied-property cache. Empty until the first Apply so adopted shapes are
' fully restyled exactly once. Geometry is deliberately absent: it diffs
' against the live shape so manual drags are corrected.
Private mAppliedValid As Boolean
Private mAppliedText As String
Private mAppliedFill As Long
Private mAppliedTextColor As Long
Private mAppliedFontSize As Double
Private mAppliedFontBold As Boolean
Private mAppliedVisible As Boolean
Private mAppliedOnAction As String
Private mAppliedValue As Double
Private mAppliedChecked As Boolean

' Async op fields.
Private mOpTask As ROneCOne
Private mOpCts As ROneCOne
Private mOpRunsProc As String
Private mOpDisableIds As Collection
Private mOpSpinnerId As String
Private mOpDoneProc As String
Private mOpFailProc As String
Private mOpCancelProc As String
Private mOpStatusKey As String
Private mOpRunning As Boolean
Private mOpLastError As String

' Job fields.
Private mJobStepProc As String
Private mJobBudgetMs As Long
Private mJobBudgetExplicit As Boolean
Private mJobPaceMs As Long
Private mJobLastStepMs As LongLong
Private mJobDoneProc As String
Private mJobFailProc As String
Private mJobCancelProc As String
Private mJobRunning As Boolean
Private mJobCancelRequested As Boolean

' Theme fields.
Private mThemePrimary As Long
Private mThemeOnPrimary As Long
Private mThemeSurface As Long
Private mThemeOnSurface As Long
Private mThemeMuted As Long
Private mThemeOnMuted As Long
Private mThemeSuccess As Long
Private mThemeDanger As Long
Private mThemeBorder As Long
Private mThemeCanvas As Long
Private mThemeFontName As String
Private mThemeFontSize As Double

' =====================================================================
' Factory surface
' =====================================================================

Public Function Mount(ByVal hostSheet As Worksheet, ByVal appId As String) As ReDimUI
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        Set Mount = ReDimUI.Mount(hostSheet, appId)
        Exit Function
    End If
    RequireIdentifier appId, False, "Mount", "appId"
    If hostSheet Is Nothing Then
        RaiseError ERROR_INVALID_ARGUMENT, "Mount", "Mount requires a worksheet."
    End If
    EnsureFactory
    If mExcelApp Is Nothing Then Set mExcelApp = Application
    Set appValue = FindApp(appId)
    If appValue Is Nothing Then
        Set appValue = New ReDimUI
        appValue.ConfigureApp appId, hostSheet
        mApps.Add appValue, appId
    Else
        appValue.RebindSheet hostSheet
    End If
    Set Mount = appValue
End Function

' Cell edits reach cell-backed TextInput components here, and closing the
' hosting workbook tears the pump down before its TIMERPROC goes stale.
Private Sub mExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim appValue As ReDimUI

    On Error Resume Next
    For Each appValue In mApps
        If appValue.Sheet Is Sh Then appValue.HandleSheetChange Target
    Next appValue
    On Error GoTo 0
End Sub

Private Sub mExcelApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Wb Is ThisWorkbook Then Shutdown
End Sub

Private Function FindApp(ByVal appId As String) As ReDimUI
    On Error Resume Next
    Set FindApp = mApps.Item(appId)
    On Error GoTo 0
End Function

Public Property Get App(ByVal appId As String) As ReDimUI
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        Set App = ReDimUI.App(appId)
        Exit Property
    End If
    EnsureFactory
    If Not TryGetObject(mApps, appId, appValue) Then
        RaiseError ERROR_NOT_FOUND, "App", _
            "No mounted app is named '" & appId & "'. Call Mount first."
    End If
    Set App = appValue
End Property

Public Property Get HasApp(ByVal appId As String) As Boolean
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        HasApp = ReDimUI.HasApp(appId)
        Exit Property
    End If
    EnsureFactory
    HasApp = TryGetObject(mApps, appId, appValue)
End Property

Public Property Get Sender() As ReDimUI
    Set Sender = ReDimUI.FactorySender
End Property

Public Property Get SenderApp() As ReDimUI
    Set SenderApp = ReDimUI.FactorySenderApp
End Property

Public Property Get SenderId() As String
    If Not ReDimUI.FactorySender Is Nothing Then
        SenderId = ReDimUI.FactorySender.ComponentId
    End If
End Property

Friend Property Get FactorySender() As ReDimUI
    Set FactorySender = mSenderComponent
End Property

Friend Property Get FactorySenderApp() As ReDimUI
    Set FactorySenderApp = mSenderApp
End Property

' Entry point for ReDimHost.RdxDispatch. Resolves the clicked shape to its
' component, applies click guards, and invokes the handler with sender context.
Friend Sub DispatchShape(ByVal shapeName As String)
    Dim appValue As ReDimUI
    Dim componentValue As ReDimUI
    Dim targetAppId As String
    Dim targetComponentId As String
    Dim targetPart As String

    If Not IsFactory(Me) Then
        ReDimUI.DispatchShape shapeName
        Exit Sub
    End If
    EnsureFactory
    If Not ParseShapeName(shapeName, targetAppId, targetComponentId, targetPart) Then
        Exit Sub
    End If
    If Not TryGetObject(mApps, targetAppId, appValue) Then Exit Sub
    If Not appValue.TryGetComponent(targetComponentId, componentValue) Then Exit Sub
    If Not componentValue.CanClick Then Exit Sub
    componentValue.MarkClicked
    Set mSenderApp = appValue
    Set mSenderComponent = componentValue
    mSenderPartName = targetPart
    On Error GoTo HandlerFailed
    componentValue.RunClickBehavior
    On Error GoTo 0
    Set mSenderApp = Nothing
    Set mSenderComponent = Nothing
    mSenderPartName = vbNullString
    Exit Sub

HandlerFailed:
    Dim failure As String
    failure = "ReDim handler error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
    Set mSenderApp = Nothing
    Set mSenderComponent = Nothing
    mSenderPartName = vbNullString
    appValue.ReportAppError failure
End Sub

Public Property Get SenderPart() As String
    SenderPart = ReDimUI.FactorySenderPart
End Property

' ---------------------------------------------------------------
' Keyboard focus: one field at a time owns typed characters, routed
' through ReDimHost.RdxKeyChar because OnKey can only target a standard
' module. Keys are bound only while a field is focused, so ordinary sheet
' typing is untouched the rest of the time.
' ---------------------------------------------------------------

Friend Sub DispatchKey(ByVal keyText As String)
    Dim appValue As ReDimUI
    Dim componentValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.DispatchKey keyText
        Exit Sub
    End If
    EnsureFactory
    If LenB(mFocusComponentId) = 0 Then Exit Sub
    Set appValue = FindApp(mFocusAppId)
    If appValue Is Nothing Then
        ClearKeyboardFocus
        Exit Sub
    End If
    Set componentValue = appValue.PeekComponent(mFocusComponentId)
    If componentValue Is Nothing Then
        ClearKeyboardFocus
        Exit Sub
    End If
    On Error Resume Next
    componentValue.HandleKey appValue, keyText
    On Error GoTo 0
End Sub

Friend Sub SetKeyboardFocus(ByVal hostApp As ReDimUI, ByVal componentValue As ReDimUI)
    Dim previousApp As ReDimUI
    Dim previousComponent As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.SetKeyboardFocus hostApp, componentValue
        Exit Sub
    End If
    EnsureFactory
    If LenB(mFocusComponentId) > 0 Then
        Set previousApp = FindApp(mFocusAppId)
        If Not previousApp Is Nothing Then
            Set previousComponent = _
                previousApp.PeekComponent(mFocusComponentId)
            If Not previousComponent Is Nothing Then
                previousComponent.BlurField previousApp, True
            End If
        End If
    End If
    mFocusAppId = hostApp.AppId
    mFocusComponentId = componentValue.ComponentId
    RdxBindKeys
    RequestPump
End Sub

Friend Sub ClearKeyboardFocus()
    If Not IsFactory(Me) Then
        ReDimUI.ClearKeyboardFocus
        Exit Sub
    End If
    mFocusAppId = vbNullString
    mFocusComponentId = vbNullString
    RdxReleaseKeys
End Sub

Public Property Get HasKeyboardFocus() As Boolean
    If Not IsFactory(Me) Then
        HasKeyboardFocus = ReDimUI.HasKeyboardFocus
        Exit Property
    End If
    HasKeyboardFocus = (LenB(mFocusComponentId) > 0)
End Property

Public Property Get FocusedComponentId() As String
    If Not IsFactory(Me) Then
        FocusedComponentId = ReDimUI.FocusedComponentId
        Exit Property
    End If
    FocusedComponentId = mFocusComponentId
End Property

Friend Function IsComponentFocused( _
    ByVal appId As String, _
    ByVal componentId As String _
) As Boolean
    If Not IsFactory(Me) Then
        IsComponentFocused = ReDimUI.IsComponentFocused(appId, componentId)
        Exit Function
    End If
    IsComponentFocused = (mFocusAppId = appId _
        And mFocusComponentId = componentId)
End Function

Friend Property Get FactorySenderPart() As String
    FactorySenderPart = mSenderPartName
End Property

' One pump tick across every mounted app. Called by ReDimHost on WM_TIMER
' (animation frame rate) and by PumpOnce in deterministic tests. Must never
' raise. Animations advance every frame with measured elapsed time; ops,
' budget jobs, and toast expiry run on a 50 ms work cadence via an
' accumulator. A forced tick is one nominal 50 ms frame with work included,
' which keeps PumpOnce-driven tests deterministic.
Friend Sub TickAll(Optional ByVal forceWork As Boolean = False)
    Dim appValue As ReDimUI
    Dim nowMs As LongLong
    Dim elapsedMs As Long
    Dim runWork As Boolean

    If Not IsFactory(Me) Then
        ReDimUI.TickAll forceWork
        Exit Sub
    End If
    EnsureFactory
    ' Frame deltas come from QueryPerformanceCounter: GetTickCount64 stays
    ' at the ~15.6 ms system tick on modern Windows even when the timer
    ' period is raised, which is too coarse for frame arithmetic. The
    ' coarse clock remains nowMs for wall-clock gates (paces, TTLs).
    Dim nowPreciseMs As Double
    nowMs = GetTickCount64()
    nowPreciseMs = PreciseNowMs()
    If forceWork Then
        elapsedMs = 50
        runWork = True
        mWorkDebtMs = 0
    Else
        If mLastTickPreciseMs = 0 Then
            elapsedMs = 16
        Else
            elapsedMs = CLng(nowPreciseMs - mLastTickPreciseMs)
        End If
        If elapsedMs < 1 Then elapsedMs = 1
        If elapsedMs > 250 Then elapsedMs = 250
        mWorkDebtMs = mWorkDebtMs + elapsedMs
        If mWorkDebtMs >= 50 Then
            runWork = True
            mWorkDebtMs = 0
        End If
    End If
    mLastTickPreciseMs = nowPreciseMs
    ' Renders inside a tick skip the ScreenUpdating toggle: flipping it at
    ' pump frequency redraws scroll bars and blips the cursor.
    mInPumpTick = True
    On Error Resume Next
    For Each appValue In mApps
        appValue.TickApp nowMs, elapsedMs, runWork
    Next appValue
    On Error GoTo 0
    mInPumpTick = False
End Sub

Friend Property Get InPumpTick() As Boolean
    If Not IsFactory(Me) Then
        InPumpTick = ReDimUI.InPumpTick
    Else
        InPumpTick = mInPumpTick
    End If
End Property

Public Sub PumpOnce()
    ReDimUI.TickAll True
End Sub

' Opt-in: pin the cursor to a steady arrow while the pump is armed. Off by
' default so interactive shapes keep their hover hand; turn on if busy
' cursor flicker is visible on your hardware.
Public Sub PinPumpCursor(ByVal enabled As Boolean)
    RdxSetCursorPin enabled
End Sub

' Deterministic tests turn the wall-clock timer off and drive PumpOnce
' themselves. Production leaves it on.
Public Sub AutoPump(ByVal enabled As Boolean)
    If Not IsFactory(Me) Then
        ReDimUI.AutoPump enabled
        Exit Sub
    End If
    mAutoPumpOff = Not enabled
End Sub

Friend Sub RequestPump()
    If Not IsFactory(Me) Then
        ReDimUI.RequestPump
        Exit Sub
    End If
    If Not mAutoPumpOff Then RdxEnsurePump
End Sub

' True while any app still owns work the pump must drive.
Friend Function HasPendingWork() As Boolean
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        HasPendingWork = ReDimUI.HasPendingWork
        Exit Function
    End If
    EnsureFactory
    For Each appValue In mApps
        If appValue.HasAppWork Then
            HasPendingWork = True
            Exit Function
        End If
    Next appValue
End Function

Public Sub Shutdown()
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.Shutdown
        Exit Sub
    End If
    EnsureFactory
    RdxStopPump
    ClearKeyboardFocus
    On Error Resume Next
    For Each appValue In mApps
        appValue.ReleaseAppResources
    Next appValue
    On Error GoTo 0
    Set mApps = New Collection
    Set mSenderApp = Nothing
    Set mSenderComponent = Nothing
End Sub

Public Property Get Version() As String
    Version = REDIM_VERSION
End Property

' ---------------------------------------------------------------
' Windows: sheets as forms, with navigation between them
' ---------------------------------------------------------------

' Shows the target window's sheet, then very-hides every other registered
' window sheet so app surfaces cannot be unhidden from the tab bar. Target
' first: Excel requires one visible sheet at all times. Ordinary sheets
' that are not registered windows are never touched. The previously active
' window goes on the back stack for NavigateBack.
Public Sub Navigate(ByVal appId As String)
    If Not IsFactory(Me) Then
        ReDimUI.Navigate appId
        Exit Sub
    End If
    NavigateCore appId, True
End Sub

Public Function NavigateBack() As Boolean
    Dim previousId As String

    If Not IsFactory(Me) Then
        NavigateBack = ReDimUI.NavigateBack
        Exit Function
    End If
    EnsureFactory
    If mNavStack Is Nothing Then Exit Function
    If mNavStack.Count = 0 Then Exit Function
    previousId = CStr(mNavStack.Item(mNavStack.Count))
    mNavStack.Remove mNavStack.Count
    NavigateCore previousId, False
    NavigateBack = True
End Function

Public Property Get ActiveWindowId() As String
    If Not IsFactory(Me) Then
        ActiveWindowId = ReDimUI.ActiveWindowId
        Exit Property
    End If
    ActiveWindowId = mActiveWindowId
End Property

Private Sub NavigateCore(ByVal appId As String, ByVal pushCurrent As Boolean)
    Dim target As ReDimUI
    Dim previous As ReDimUI
    Dim windowId As Variant
    Dim windowApp As ReDimUI
    Dim previousId As String

    EnsureFactory
    Set target = App(appId)
    If Not target.IsWindow Then
        RaiseError ERROR_INVALID_OPERATION, "Navigate", _
            "App '" & appId & "' is not a window. Call AsWindow on it first."
    End If
    If mActiveWindowId = appId Then Exit Sub
    previousId = mActiveWindowId
    target.Sheet.Visible = xlSheetVisible
    target.Sheet.Activate
    target.ReassertCanvas
    For Each windowId In mWindows
        If CStr(windowId) <> appId Then
            If TryGetObject(mApps, CStr(windowId), windowApp) Then
                On Error Resume Next
                windowApp.Sheet.Visible = xlSheetVeryHidden
                On Error GoTo 0
            End If
        End If
    Next windowId
    If LenB(previousId) > 0 Then
        If pushCurrent Then
            If mNavStack Is Nothing Then Set mNavStack = New Collection
            mNavStack.Add previousId
        End If
        If TryGetObject(mApps, previousId, previous) Then previous.InvokeHide
    End If
    mActiveWindowId = appId
    For Each windowId In mWindows
        If TryGetObject(mApps, CStr(windowId), windowApp) Then
            windowApp.RefreshNavBar
        End If
    Next windowId
    ' The newly shown window may carry pump demand (sliders, spinners).
    RequestPump
    target.InvokeShow
End Sub

Friend Function WindowIds() As Collection
    Dim windowId As Variant

    If Not IsFactory(Me) Then
        Set WindowIds = ReDimUI.WindowIds
        Exit Function
    End If
    EnsureFactory
    Set WindowIds = New Collection
    If mWindows Is Nothing Then Exit Function
    For Each windowId In mWindows
        WindowIds.Add CStr(windowId)
    Next windowId
End Function

Friend Sub RegisterWindow(ByVal appId As String)
    Dim windowId As Variant

    If Not IsFactory(Me) Then
        ReDimUI.RegisterWindow appId
        Exit Sub
    End If
    EnsureFactory
    If mWindows Is Nothing Then Set mWindows = New Collection
    For Each windowId In mWindows
        If CStr(windowId) = appId Then Exit Sub
    Next windowId
    mWindows.Add appId
End Sub

Friend Sub ForgetWindow(ByVal appId As String)
    Dim position As Long

    If Not IsFactory(Me) Then
        ReDimUI.ForgetWindow appId
        Exit Sub
    End If
    EnsureFactory
    If mWindows Is Nothing Then Exit Sub
    For position = mWindows.Count To 1 Step -1
        If CStr(mWindows.Item(position)) = appId Then mWindows.Remove position
    Next position
    If Not mNavStack Is Nothing Then
        For position = mNavStack.Count To 1 Step -1
            If CStr(mNavStack.Item(position)) = appId Then
                mNavStack.Remove position
            End If
        Next position
    End If
    If mActiveWindowId = appId Then mActiveWindowId = vbNullString
End Sub

Public Function ThemeLight() As ReDimUI
    Dim themeValue As ReDimUI

    Set themeValue = New ReDimUI
    themeValue.ConfigureTheme _
        RGB(33, 115, 70), RGB(255, 255, 255), _
        RGB(255, 255, 255), RGB(50, 49, 48), _
        RGB(237, 235, 233), RGB(120, 118, 116), _
        RGB(16, 124, 16), RGB(196, 43, 28), _
        RGB(225, 223, 221), RGB(250, 249, 248), _
        "Segoe UI", 11
    Set ThemeLight = themeValue
End Function

Public Function ThemeDark() As ReDimUI
    Dim themeValue As ReDimUI

    Set themeValue = New ReDimUI
    themeValue.ConfigureTheme _
        RGB(70, 160, 105), RGB(24, 24, 24), _
        RGB(50, 49, 48), RGB(243, 242, 241), _
        RGB(72, 70, 68), RGB(180, 178, 176), _
        RGB(108, 203, 95), RGB(241, 112, 123), _
        RGB(96, 94, 92), RGB(32, 31, 30), _
        "Segoe UI", 11
    Set ThemeDark = themeValue
End Function

' =====================================================================
' App configuration and surface
' =====================================================================

Friend Sub ConfigureApp(ByVal appId As String, ByVal hostSheet As Worksheet)
    mRole = ROLE_APP
    mAppId = appId
    Set mSheet = hostSheet
    Set mComponents = New Collection
    Set mComponentOrder = New Collection
    Set mStateValues = New Collection
    Set mStateHandlers = New Collection
    Set mDirtyIds = New Collection
    Set mAsyncOps = New Collection
    Set mJobs = New Collection
    Set mTheme = ReDimUI.ThemeLight
End Sub

Friend Sub RebindSheet(ByVal hostSheet As Worksheet)
    RequireRole ROLE_APP, "RebindSheet"
    Set mSheet = hostSheet
End Sub

Public Property Get AppId() As String
    RequireRole ROLE_APP, "AppId"
    AppId = mAppId
End Property

Public Property Get Sheet() As Worksheet
    RequireRole ROLE_APP, "Sheet"
    Set Sheet = mSheet
End Property

Public Function Button(ByVal componentId As String) As ReDimUI
    Set Button = GetOrCreateComponent(componentId, KIND_BUTTON)
End Function

Public Function Label(ByVal componentId As String) As ReDimUI
    Set Label = GetOrCreateComponent(componentId, KIND_LABEL)
End Function

Public Function Card(ByVal componentId As String) As ReDimUI
    Set Card = GetOrCreateComponent(componentId, KIND_CARD)
End Function

Public Function Spinner(ByVal componentId As String) As ReDimUI
    Set Spinner = GetOrCreateComponent(componentId, KIND_SPINNER)
End Function

Public Function ProgressBar(ByVal componentId As String) As ReDimUI
    Set ProgressBar = GetOrCreateComponent(componentId, KIND_PROGRESS)
End Function

Public Function Toggle(ByVal componentId As String) As ReDimUI
    Set Toggle = GetOrCreateComponent(componentId, KIND_TOGGLE)
End Function

' Drawn picker: face, caret, and option list in the app font.
Public Function SelectBox(ByVal componentId As String) As ReDimUI
    Set SelectBox = GetOrCreateComponent(componentId, KIND_SELECT)
End Function

' Drawn checkbox: themed box, check glyph, and caption. Click the box or
' the caption to toggle.
Public Function TickBox(ByVal componentId As String) As ReDimUI
    Set TickBox = GetOrCreateComponent(componentId, KIND_TICKBOX)
End Function

' Drawn single-select option rows. Items supplies the rows, Value the
' selected index, and WritesTo receives the selected item's text.
Public Function RadioGroup(ByVal componentId As String) As ReDimUI
    Set RadioGroup = GetOrCreateComponent(componentId, KIND_RADIO)
End Function

' Drawn numeric control: minus and plus around a value face, honoring
' SliderRange. Shapes cannot report click coordinates, so a drag-thumb
' slider cannot be drawn honestly; stepped entry can.
Public Function Stepper(ByVal componentId As String) As ReDimUI
    Set Stepper = GetOrCreateComponent(componentId, KIND_STEPPER)
End Function

' Drawn slider, click-to-set: the click position is mapped back into sheet
' points by inverting PointsToScreenPixels, then snapped to SliderRange.
' Loop-free by design; nothing in ReDim may block the pump.
Public Function SlideBar(ByVal componentId As String) As ReDimUI
    Set SlideBar = GetOrCreateComponent(componentId, KIND_SLIDEBAR)
End Function

' Editable combo: a cell-backed text face with a caret and a drawn option
' list. Excel pauses all VBA during cell edit mode, so filtering applies
' at the moments the platform grants: opening the list filters by the
' typed text, and committing with Enter auto-suggests matches (an exact
' match simply takes it). Free text is always allowed.
Public Function ComboBox(ByVal componentId As String) As ReDimUI
    Set ComboBox = GetOrCreateComponent(componentId, KIND_COMBO)
End Function

Public Function TextInput(ByVal componentId As String) As ReDimUI
    Set TextInput = GetOrCreateComponent(componentId, KIND_INPUT)
End Function

' Transient notification card. Expires on the pump after ttlMs, or when
' clicked.
Public Function Toast( _
    ByVal message As String, _
    Optional ByVal ttlMs As Long = 3000 _
) As ReDimUI
    Dim toastValue As ReDimUI

    RequireRole ROLE_APP, "Toast"
    mToastCounter = mToastCounter + 1
    Set toastValue = GetOrCreateComponent("toast_" & mToastCounter, KIND_TOAST)
    toastValue.ConfigureToast message, ttlMs, NextToastSlot()
    NotifyDirty toastValue.ComponentId
    If mRendered Then FlushDirty
    ReDimUI.RequestPump
    Set Toast = toastValue
End Function

' Pins the toast tray's top-left to a range anchor instead of the default
' rail beside the app content.
Public Function ToastTray(ByVal anchor As String) As ReDimUI
    RequireRole ROLE_APP, "ToastTray"
    mToastTrayAnchor = anchor
    Set ToastTray = Me
End Function

Friend Function ToastTrayOrigin( _
    ByRef originLeft As Double, _
    ByRef originTop As Double _
) As Boolean
    Dim anchorRange As Range

    If LenB(mToastTrayAnchor) = 0 Then Exit Function
    On Error Resume Next
    Set anchorRange = mSheet.Range(mToastTrayAnchor)
    On Error GoTo 0
    If anchorRange Is Nothing Then Exit Function
    originLeft = anchorRange.Left
    originTop = anchorRange.Top
    ToastTrayOrigin = True
End Function

' Lowest slot not held by a live toast, so slot one is reused the moment the
' previous toast expires instead of the stack growing downward forever.
Private Function NextToastSlot() As Long
    Dim componentValue As ReDimUI
    Dim candidate As Long
    Dim taken As Boolean

    For candidate = 1 To mComponents.Count + 1
        taken = False
        For Each componentValue In mComponents
            If componentValue.ToastSlot = candidate Then
                taken = True
                Exit For
            End If
        Next componentValue
        If Not taken Then
            NextToastSlot = candidate
            Exit Function
        End If
    Next candidate
    NextToastSlot = 1
End Function

' Shapes-only modal: a click-swallowing dim overlay plus a centered card
' with OK and optional Cancel buttons. No UserForms.
Public Sub Confirm( _
    ByVal title As String, _
    ByVal message As String, _
    Optional ByVal okProc As String, _
    Optional ByVal cancelProc As String, _
    Optional ByVal okText As String = "OK", _
    Optional ByVal cancelText As String = "Cancel" _
)
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim cardLeft As Double
    Dim cardTop As Double

    RequireRole ROLE_APP, "Confirm"
    ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    cardLeft = viewLeft + (viewWidth - 340) / 2
    cardTop = viewTop + (viewHeight - 150) / 2
    BeginUpdate
    ' The dim sheet runs from the origin to one viewport past the visible
    ' area, so ordinary scrolling while the modal is up stays covered.
    GetOrCreateComponent("mdl_ov", KIND_OVERLAY) _
        .AtRect(0, 0, viewLeft + viewWidth * 2, viewTop + viewHeight * 2) _
        .Visible True
    Card("mdl_card").AtRect(cardLeft, cardTop, 340, 150) _
        .Text(title & vbLf & vbLf & message).Visible True
    Button("mdl_ok").AtRect(cardLeft + 340 - 200, cardTop + 150 - 44, 90, 30) _
        .Text(okText).Primary.Visible True
    Button("mdl_ok").OnClick okProc
    Component("mdl_ok").MarkCloseModalOnClick
    Button("mdl_cancel").AtRect(cardLeft + 340 - 102, cardTop + 150 - 44, 90, 30) _
        .Text(cancelText).Secondary.Visible(LenB(cancelText) > 0)
    Button("mdl_cancel").OnClick cancelProc
    Component("mdl_cancel").MarkCloseModalOnClick
    EndUpdate
    If Not mRendered Then Render
    RaiseModalZOrder
End Sub

Public Sub CloseModal()
    RequireRole ROLE_APP, "CloseModal"
    BeginUpdate
    HideIfExists "mdl_ov"
    HideIfExists "mdl_card"
    HideIfExists "mdl_ok"
    HideIfExists "mdl_cancel"
    EndUpdate
End Sub

Private Sub HideIfExists(ByVal componentId As String)
    Dim componentValue As ReDimUI

    If TryGetComponent(componentId, componentValue) Then
        componentValue.Visible False
    End If
End Sub

Private Sub RaiseModalZOrder()
    Dim componentId As Variant
    Dim componentValue As ReDimUI

    For Each componentId In Array("mdl_ov", "mdl_card", "mdl_ok", "mdl_cancel")
        If TryGetComponent(CStr(componentId), componentValue) Then
            On Error Resume Next
            mSheet.Shapes(componentValue.ShapeName).ZOrder msoBringToFront
            On Error GoTo 0
        End If
    Next componentId
End Sub

Friend Sub ResolveViewport( _
    ByRef viewLeft As Double, _
    ByRef viewTop As Double, _
    ByRef viewWidth As Double, _
    ByRef viewHeight As Double _
)
    Dim visible As Range

    On Error Resume Next
    If ActiveWindow Is Nothing Then
    ElseIf mSheet Is ActiveWindow.ActiveSheet Then
        Set visible = ActiveWindow.VisibleRange
    End If
    On Error GoTo 0
    If visible Is Nothing Then Set visible = mSheet.Range("A1:P36")
    viewLeft = visible.Left
    viewTop = visible.Top
    viewWidth = visible.Width
    viewHeight = visible.Height
End Sub

' Paints the sheet background and hides gridlines so shapes sit on a clean
' canvas. Gridlines are window state, so they change only while this sheet
' is active.
Public Function PrepareCanvas( _
    Optional ByVal hideGridlines As Boolean = True, _
    Optional ByVal paintBackground As Boolean = True _
) As ReDimUI
    RequireRole ROLE_APP, "PrepareCanvas"
    If paintBackground Then mSheet.Cells.Interior.Color = mTheme.CanvasColor
    If hideGridlines Then
        mCanvasPrepared = True
        On Error Resume Next
        If mSheet Is ActiveWindow.ActiveSheet Then
            ActiveWindow.DisplayGridlines = False
        End If
        On Error GoTo 0
    End If
    Set PrepareCanvas = Me
End Function

Friend Sub HandleSheetChange(ByVal changed As Range)
    Dim componentValue As ReDimUI

    For Each componentValue In mComponents
        componentValue.SyncInputFromSheet Me, changed
    Next componentValue
End Sub

' The rectangle the app's durable components occupy. Toasts anchor here so
' they appear beside the UI, not at the window's far edge on wide monitors.
Friend Function ContentBounds( _
    ByRef boundLeft As Double, _
    ByRef boundTop As Double, _
    ByRef boundRight As Double _
) As Boolean
    Dim componentValue As ReDimUI

    For Each componentValue In mComponents
        If componentValue.CountsForBounds Then
            If Not ContentBounds Then
                boundLeft = componentValue.BoundsLeft
                boundTop = componentValue.BoundsTop
                boundRight = componentValue.BoundsRight
                ContentBounds = True
            Else
                If componentValue.BoundsLeft < boundLeft Then
                    boundLeft = componentValue.BoundsLeft
                End If
                If componentValue.BoundsTop < boundTop Then
                    boundTop = componentValue.BoundsTop
                End If
                If componentValue.BoundsRight > boundRight Then
                    boundRight = componentValue.BoundsRight
                End If
            End If
        End If
    Next componentValue
End Function

Friend Function GetOrCreateComponent( _
    ByVal componentId As String, _
    ByVal componentKind As Long _
) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Component"
    RequireIdentifier componentId, True, "Component", "componentId"
    If TryGetComponent(componentId, componentValue) Then
        If componentValue.ComponentKind <> componentKind Then
            RaiseError ERROR_INVALID_OPERATION, "Component", _
                "Component '" & componentId & "' already exists with a different kind."
        End If
        Set GetOrCreateComponent = componentValue
        Exit Function
    End If
    Set componentValue = New ReDimUI
    componentValue.ConfigureComponent mAppId, componentId, componentKind
    mComponents.Add componentValue, componentId
    mComponentOrder.Add componentId
    Set GetOrCreateComponent = componentValue
End Function

Friend Function TryGetComponent( _
    ByVal componentId As String, _
    ByRef result As ReDimUI _
) As Boolean
    TryGetComponent = TryGetObject(mComponents, componentId, result)
End Function

Friend Function PeekComponent(ByVal componentId As String) As ReDimUI
    Set PeekComponent = FindKeyed(mComponents, componentId)
End Function

Public Property Get Component(ByVal componentId As String) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Component"
    If Not TryGetComponent(componentId, componentValue) Then
        RaiseError ERROR_NOT_FOUND, "Component", _
            "No component is named '" & componentId & "'."
    End If
    Set Component = componentValue
End Property

Public Property Get ComponentCount() As Long
    RequireRole ROLE_APP, "ComponentCount"
    ComponentCount = mComponents.Count
End Property

' ---------------------------------------------------------------
' State store
' ---------------------------------------------------------------

Public Sub SetState(ByVal stateKey As String, ByVal stateValue As Variant)
    RequireRole ROLE_APP, "SetState"
    ReplaceItem mStateValues, stateKey, stateValue
    If mPersistOn Then WritePersistedValue stateKey, stateValue
    MarkBoundComponentsDirty stateKey
    InvokeStateHandlers stateKey
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

' Sets the key only when it has no value yet. Setup code uses this for
' initial values so a persisted store keeps the user's last choices.
Public Sub SetStateDefault(ByVal stateKey As String, ByVal stateValue As Variant)
    RequireRole ROLE_APP, "SetStateDefault"
    If HasState(stateKey) Then Exit Sub
    SetState stateKey, stateValue
End Sub

Public Property Get State(ByVal stateKey As String) As Variant
    Dim stored As Variant

    RequireRole ROLE_APP, "State"
    If Not TryGetItem(mStateValues, stateKey, stored) Then
        RaiseError ERROR_NOT_FOUND, "State", _
            "No state value is named '" & stateKey & "'. Set it with SetState first."
    End If
    AssignVariant State, stored
End Property

Public Property Get StateOrDefault( _
    ByVal stateKey As String, _
    ByVal fallback As Variant _
) As Variant
    Dim stored As Variant

    RequireRole ROLE_APP, "StateOrDefault"
    If TryGetItem(mStateValues, stateKey, stored) Then
        AssignVariant StateOrDefault, stored
    Else
        AssignVariant StateOrDefault, fallback
    End If
End Property

Public Property Get HasState(ByVal stateKey As String) As Boolean
    Dim stored As Variant

    RequireRole ROLE_APP, "HasState"
    HasState = TryGetItem(mStateValues, stateKey, stored)
End Property

' Persistence writes state through to hidden workbook names, so a reopened
' workbook resumes where the user left off. Turning it on hydrates the store
' from any values a previous session persisted; call it before Render and
' use SetStateDefault for initial values.
Public Function Persist(ByVal enabled As Boolean) As ReDimUI
    RequireRole ROLE_APP, "Persist"
    mPersistOn = enabled
    If enabled Then HydratePersistedState
    Set Persist = Me
End Function

Public Sub ClearPersisted()
    Dim book As Workbook
    Dim position As Long
    Dim prefix As String

    RequireRole ROLE_APP, "ClearPersisted"
    Set book = mSheet.Parent
    prefix = PERSIST_PREFIX & mAppId & "_"
    For position = book.Names.Count To 1 Step -1
        If Left$(book.Names(position).Name, Len(prefix)) = prefix Then
            book.Names(position).Delete
        End If
    Next position
End Sub

Private Sub HydratePersistedState()
    Dim book As Workbook
    Dim entry As Name
    Dim prefix As String
    Dim stateKey As String
    Dim decoded As Variant

    Set book = mSheet.Parent
    prefix = PERSIST_PREFIX & mAppId & "_"
    For Each entry In book.Names
        If Left$(entry.Name, Len(prefix)) = prefix Then
            stateKey = Mid$(entry.Name, Len(prefix) + 1)
            If TryDecodePersisted(entry.RefersTo, decoded) Then
                ReplaceItem mStateValues, stateKey, decoded
                MarkBoundComponentsDirty stateKey
            End If
        End If
    Next entry
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

Private Sub WritePersistedValue(ByVal stateKey As String, ByVal stateValue As Variant)
    Dim encoded As String

    If Not TryEncodePersisted(stateValue, encoded) Then Exit Sub
    On Error Resume Next
    mSheet.Parent.Names.Add PERSIST_PREFIX & mAppId & "_" & stateKey, _
        "=""" & Replace(encoded, """", """""") & """", False
    On Error GoTo 0
End Sub

' Scalars persist with a one-letter type tag. Objects and arrays are
' in-memory only and are skipped silently.
Private Function TryEncodePersisted( _
    ByVal stateValue As Variant, _
    ByRef encoded As String _
) As Boolean
    If IsObject(stateValue) Or IsArray(stateValue) Then Exit Function
    Select Case VarType(stateValue)
        Case vbBoolean
            encoded = "B:" & CStr(CBool(stateValue))
        Case vbByte, vbInteger, vbLong, vbLongLong
            encoded = "L:" & CStr(stateValue)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            encoded = "D:" & CStr(CDbl(stateValue))
        Case vbDate
            encoded = "T:" & CStr(CDbl(CDate(stateValue)))
        Case vbString
            encoded = "S:" & stateValue
        Case Else
            Exit Function
    End Select
    TryEncodePersisted = True
End Function

Private Function TryDecodePersisted( _
    ByVal storedRefersTo As String, _
    ByRef decoded As Variant _
) As Boolean
    Dim body As String

    If Left$(storedRefersTo, 2) <> "=""" Then Exit Function
    If Right$(storedRefersTo, 1) <> """" Then Exit Function
    body = Mid$(storedRefersTo, 3, Len(storedRefersTo) - 3)
    body = Replace(body, """""", """")
    If Len(body) < 2 Then Exit Function
    On Error GoTo BadValue
    Select Case Left$(body, 2)
        Case "B:"
            decoded = CBool(Mid$(body, 3))
        Case "L:"
            decoded = CLng(Mid$(body, 3))
        Case "D:"
            decoded = CDbl(Mid$(body, 3))
        Case "T:"
            decoded = CDate(CDbl(Mid$(body, 3)))
        Case "S:"
            decoded = Mid$(body, 3)
        Case Else
            Exit Function
    End Select
    TryDecodePersisted = True
BadValue:
End Function

' Registers a zero-argument workbook procedure invoked after the key changes.
Public Sub OnStateChanged(ByVal stateKey As String, ByVal procName As String)
    Dim handlers As Collection

    RequireRole ROLE_APP, "OnStateChanged"
    Set handlers = FindHandlers(stateKey)
    If handlers Is Nothing Then
        Set handlers = New Collection
        mStateHandlers.Add handlers, stateKey
    End If
    handlers.Add procName
End Sub

Private Sub InvokeStateHandlers(ByVal stateKey As String)
    Dim handlers As Collection
    Dim procName As Variant

    Set handlers = FindHandlers(stateKey)
    If handlers Is Nothing Then Exit Sub
    For Each procName In handlers
        Application.Run CStr(procName)
    Next procName
End Sub

' TryGetObject is typed for ReDimUI values; handler lists are Collections
' and need their own lookup or the ByRef assignment fails at run time and
' reads as "not found".
Private Function FindHandlers(ByVal stateKey As String) As Collection
    On Error Resume Next
    Set FindHandlers = mStateHandlers.Item(stateKey)
    On Error GoTo 0
End Function

Private Sub MarkBoundComponentsDirty(ByVal stateKey As String)
    Dim componentValue As ReDimUI

    For Each componentValue In mComponents
        If componentValue.IsBoundTo(stateKey) Then
            NotifyDirty componentValue.ComponentId
        End If
    Next componentValue
End Sub

' ---------------------------------------------------------------
' Rendering
' ---------------------------------------------------------------

Public Sub BeginUpdate()
    RequireRole ROLE_APP, "BeginUpdate"
    mBatchDepth = mBatchDepth + 1
End Sub

Public Sub EndUpdate()
    RequireRole ROLE_APP, "EndUpdate"
    If mBatchDepth > 0 Then mBatchDepth = mBatchDepth - 1
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

Public Sub Render()
    Dim componentId As Variant

    RequireRole ROLE_APP, "Render"
    For Each componentId In mComponentOrder
        NotifyDirty CStr(componentId)
    Next componentId
    mRendered = True
    FlushDirty
    PruneOrphanShapes
    ' Anything rendered that needs frames (a slider's press watch, a
    ' spinner, a toast) arms the pump right here; demand alone only keeps
    ' an already-armed pump alive.
    If HasAppWork() Then ReDimUI.RequestPump
End Sub

' Shapes with this app's prefix but no registered component are leftovers
' from renamed or removed components in an earlier build. Render sweeps them
' so iterating on app code never litters the sheet.
Private Sub PruneOrphanShapes()
    Dim target As Shape
    Dim orphanNames As Collection
    Dim orphanName As Variant
    Dim prefix As String
    Dim remainder As String
    Dim orphanComponentId As String
    Dim partAt As Long
    Dim componentValue As ReDimUI

    Set orphanNames = New Collection
    prefix = SHAPE_PREFIX & mAppId & "_"
    For Each target In mSheet.Shapes
        If Left$(target.Name, Len(prefix)) = prefix Then
            remainder = Mid$(target.Name, Len(prefix) + 1)
            partAt = InStr(1, remainder, PART_SEPARATOR)
            If partAt > 0 Then
                orphanComponentId = Left$(remainder, partAt - 1)
            Else
                orphanComponentId = remainder
            End If
            If Not TryGetComponent(orphanComponentId, componentValue) Then
                orphanNames.Add target.Name
            End If
        End If
    Next target
    On Error Resume Next
    For Each orphanName In orphanNames
        mSheet.Shapes(CStr(orphanName)).Delete
    Next orphanName
    On Error GoTo 0
End Sub

Friend Sub NotifyDirty(ByVal componentId As String)
    Dim existing As Variant

    If TryGetItem(mDirtyIds, componentId, existing) Then Exit Sub
    mDirtyIds.Add componentId, componentId
End Sub

Friend Sub NotifyComponentTouched(ByVal componentId As String)
    NotifyDirty componentId
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

Public Sub FlushDirty()
    Dim componentValue As ReDimUI
    Dim componentId As Variant
    Dim priorScreenUpdating As Boolean
    Dim toggledScreen As Boolean

    RequireRole ROLE_APP, "FlushDirty"
    If mDirtyIds.Count = 0 Then Exit Sub
    priorScreenUpdating = Application.ScreenUpdating
    If mDirtyIds.Count > 1 And Not ReDimUI.InPumpTick Then
        Application.ScreenUpdating = False
        toggledScreen = True
    End If
    On Error GoTo RestoreScreen
    For Each componentId In mDirtyIds
        If TryGetComponent(CStr(componentId), componentValue) Then
            componentValue.Apply mSheet, mTheme, Me
        End If
    Next componentId
    Set mDirtyIds = New Collection
    If toggledScreen Then Application.ScreenUpdating = priorScreenUpdating
    Exit Sub

RestoreScreen:
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim capturedDescription As String
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    If toggledScreen Then Application.ScreenUpdating = priorScreenUpdating
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Sub

Public Function SetTheme(ByVal themeValue As ReDimUI) As ReDimUI
    RequireRole ROLE_APP, "SetTheme"
    If themeValue Is Nothing Then
        RaiseError ERROR_INVALID_ARGUMENT, "SetTheme", "SetTheme requires a theme."
    End If
    Set mTheme = themeValue
    Render
    Set SetTheme = Me
End Function

Public Property Get Theme() As ReDimUI
    RequireRole ROLE_APP, "Theme"
    Set Theme = mTheme
End Property

' Registers this app as a window: its sheet participates in Navigate and
' NavigateBack, shown alone among registered windows like a form.
Public Function AsWindow() As ReDimUI
    RequireRole ROLE_APP, "AsWindow"
    mIsWindow = True
    ReDimUI.RegisterWindow mAppId
    Set AsWindow = Me
End Function

Public Property Get IsWindow() As Boolean
    RequireRole ROLE_APP, "IsWindow"
    IsWindow = mIsWindow
End Property

' Display name for navigation chrome; the app id is the fallback.
Public Function WindowTitle(ByVal title As String) As ReDimUI
    RequireRole ROLE_APP, "WindowTitle"
    mWindowTitle = title
    Set WindowTitle = Me
End Function

Friend Property Get WindowTitleOrId() As String
    If LenB(mWindowTitle) > 0 Then
        WindowTitleOrId = mWindowTitle
    Else
        WindowTitleOrId = mAppId
    End If
End Property

' Renders one tab per registered window across the top of this app's
' sheet, active tab highlighted. Bars refresh on every Navigate, so late
' registrations and removed windows are picked up automatically.
Public Function NavBar( _
    Optional ByVal leftPoints As Double = 24, _
    Optional ByVal topPoints As Double = 8, _
    Optional ByVal tabWidth As Double = 110, _
    Optional ByVal tabHeight As Double = 26 _
) As ReDimUI
    RequireRole ROLE_APP, "NavBar"
    mHasNavBar = True
    mNavBarLeft = leftPoints
    mNavBarTop = topPoints
    mNavBarTabWidth = tabWidth
    mNavBarTabHeight = tabHeight
    RefreshNavBar
    Set NavBar = Me
End Function

Friend Sub RefreshNavBar()
    Dim windowIds As Collection
    Dim windowId As Variant
    Dim componentValue As ReDimUI
    Dim staleIds As Collection
    Dim staleId As Variant
    Dim known As Boolean
    Dim position As Long
    Dim tabButton As ReDimUI

    If Not mHasNavBar Then Exit Sub
    Set windowIds = ReDimUI.WindowIds
    BeginUpdate
    ' Tabs for windows that no longer exist get removed before rebuilding.
    Set staleIds = New Collection
    For Each componentValue In mComponents
        If Left$(componentValue.ComponentId, 4) = "nvb_" Then
            known = False
            For Each windowId In windowIds
                If "nvb_" & CStr(windowId) = componentValue.ComponentId Then
                    known = True
                    Exit For
                End If
            Next windowId
            If Not known Then staleIds.Add componentValue.ComponentId
        End If
    Next componentValue
    For Each staleId In staleIds
        Component(CStr(staleId)).Remove
    Next staleId
    For Each windowId In windowIds
        position = position + 1
        Set tabButton = Button("nvb_" & CStr(windowId))
        tabButton.AtRect mNavBarLeft + (position - 1) * (mNavBarTabWidth + 6), _
            mNavBarTop, mNavBarTabWidth, mNavBarTabHeight
        tabButton.Text ReDimUI.App(CStr(windowId)).WindowTitleOrId
        tabButton.NavigatesTo CStr(windowId)
        If ReDimUI.ActiveWindowId = CStr(windowId) Then
            tabButton.Primary
        Else
            tabButton.Secondary
        End If
    Next windowId
    EndUpdate
End Sub

' Zero-argument workbook procedures run after this window is shown or
' hidden by navigation, in the spirit of form Load and Deactivate.
Public Function OnShow(ByVal procName As String) As ReDimUI
    RequireRole ROLE_APP, "OnShow"
    mOnShowProc = procName
    Set OnShow = Me
End Function

Public Function OnHide(ByVal procName As String) As ReDimUI
    RequireRole ROLE_APP, "OnHide"
    mOnHideProc = procName
    Set OnHide = Me
End Function

Friend Sub InvokeShow()
    If LenB(mOnShowProc) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run mOnShowProc
    On Error GoTo 0
End Sub

Friend Sub InvokeHide()
    If LenB(mOnHideProc) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run mOnHideProc
    On Error GoTo 0
End Sub

' Gridlines are window state, not sheet state, so a freshly activated
' window re-hides them if its canvas was prepared.
Friend Sub ReassertCanvas()
    If Not mCanvasPrepared Then Exit Sub
    On Error Resume Next
    If mSheet Is ActiveWindow.ActiveSheet Then
        ActiveWindow.DisplayGridlines = False
    End If
    On Error GoTo 0
End Sub

' Treats the app sheet as an application surface: users cannot enter cell
' edit mode or drag shapes there, which also keeps the pump from stalling
' in edit mode, while UserInterfaceOnly leaves every framework write free.
' TextInput cells stay unlocked and editable. Opt-in, and scoped to this
' one sheet: the rest of the workbook is untouched. Call after Render.
Public Function ProtectSurface(Optional ByVal enabled As Boolean = True) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "ProtectSurface"
    If enabled Then
        For Each componentValue In mComponents
            componentValue.UnlockInputCell mSheet
        Next componentValue
        mSheet.Protect DrawingObjects:=True, Contents:=True, _
            UserInterfaceOnly:=True
        mSurfaceProtected = True
    Else
        On Error Resume Next
        mSheet.Unprotect
        On Error GoTo 0
        mSurfaceProtected = False
    End If
    Set ProtectSurface = Me
End Function

Public Property Get IsSurfaceProtected() As Boolean
    RequireRole ROLE_APP, "IsSurfaceProtected"
    IsSurfaceProtected = mSurfaceProtected
End Property

' Application.OnKey with lifecycle management: registered keys are released
' on Unmount and Shutdown so no binding outlives its app.
Public Sub HotKey(ByVal keyCode As String, ByVal procName As String)
    RequireRole ROLE_APP, "HotKey"
    Application.OnKey keyCode, procName
    If mHotKeys Is Nothing Then Set mHotKeys = New Collection
    mHotKeys.Add keyCode
End Sub

Public Sub ClearHotKeys()
    Dim keyCode As Variant

    RequireRole ROLE_APP, "ClearHotKeys"
    If mHotKeys Is Nothing Then Exit Sub
    On Error Resume Next
    For Each keyCode In mHotKeys
        Application.OnKey CStr(keyCode)
    Next keyCode
    On Error GoTo 0
    Set mHotKeys = New Collection
End Sub

Friend Sub ReleaseAppResources()
    RequireRole ROLE_APP, "ReleaseAppResources"
    ClearHotKeys
End Sub

' Removes every component shape and unregisters the app.
Public Sub Unmount(Optional ByVal deleteShapes As Boolean = True)
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Unmount"
    ClearHotKeys
    If ReDimUI.FocusedComponentId <> vbNullString Then
        If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
            ReDimUI.ClearKeyboardFocus
        End If
    End If
    If mSurfaceProtected Then ProtectSurface False
    If deleteShapes Then
        For Each componentValue In mComponents
            componentValue.DeleteShapes mSheet
        Next componentValue
    End If
    Set mComponents = New Collection
    Set mComponentOrder = New Collection
    Set mDirtyIds = New Collection
    ReDimUI.ForgetApp mAppId
End Sub

Friend Sub ForgetApp(ByVal appId As String)
    Dim existing As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.ForgetApp appId
        Exit Sub
    End If
    EnsureFactory
    If TryGetObject(mApps, appId, existing) Then mApps.Remove appId
    ForgetWindow appId
End Sub

Friend Sub ReportAppError(ByVal message As String)
    RequireRole ROLE_APP, "ReportAppError"
    If LenB(mErrorHandlerProc) > 0 Then
        On Error Resume Next
        Application.Run mErrorHandlerProc, message
        On Error GoTo 0
    Else
        Application.StatusBar = message
    End If
End Sub

' Registers a one-argument workbook procedure that receives handler failures.
Public Sub OnError(ByVal procName As String)
    RequireRole ROLE_APP, "OnError"
    mErrorHandlerProc = procName
End Sub

Friend Sub TickApp( _
    ByVal nowMs As LongLong, _
    ByVal elapsedMs As Long, _
    ByVal runWork As Boolean _
)
    Dim opValue As ReDimUI
    Dim jobValue As ReDimUI
    Dim componentValue As ReDimUI
    Dim expiredIds As Collection

    Dim animating As Boolean
    For Each componentValue In mComponents
        If componentValue.NeedsAnimation Then
            animating = True
            Exit For
        End If
    Next componentValue
    If runWork Then
        For Each opValue In mAsyncOps
            If opValue.OpIsRunning Then opValue.TickOp Me
        Next opValue
    End If
    ' Paced jobs self-gate on the wall clock, so sampling them every frame
    ' tightens their timing; budget jobs stay on the work cadence and give
    ' animations room while they run.
    For Each jobValue In mJobs
        If jobValue.JobIsRunning Then
            jobValue.TickJob Me, nowMs, runWork, animating
        End If
    Next jobValue
    ' Sliders drag with the left button only; the field blur watch treats
    ' a press of either button as clicking away.
    Dim leftButtonDown As Boolean
    Dim anyButtonDown As Boolean
    Dim pressEdge As Boolean
    leftButtonDown = (GetAsyncKeyState(VK_LEFT_BUTTON) < 0)
    anyButtonDown = leftButtonDown _
        Or (GetAsyncKeyState(VK_RIGHT_BUTTON) < 0)
    pressEdge = (anyButtonDown And Not mFieldPrevButtonDown)
    mFieldPrevButtonDown = anyButtonDown
    Set expiredIds = New Collection
    For Each componentValue In mComponents
        componentValue.TickAnimation mSheet, elapsedMs
        componentValue.TickSlideDrag Me, leftButtonDown
        componentValue.TickFieldWatch Me, anyButtonDown, pressEdge
        If runWork Then
            If componentValue.IsExpiredToast(nowMs) Then
                expiredIds.Add componentValue.ComponentId
            End If
        End If
    Next componentValue
    Dim expiredId As Variant
    For Each expiredId In expiredIds
        If TryGetComponent(CStr(expiredId), componentValue) Then
            componentValue.Remove
        End If
    Next expiredId
    If expiredIds.Count > 0 Then CompactToastSlots
End Sub

' After any toast leaves, survivors renumber to slots 1..N in their current
' order and slide up into the freed positions, the way a notification tray
' behaves.
Private Sub CompactToastSlots()
    Dim live As Collection
    Dim componentValue As ReDimUI
    Dim lowest As ReDimUI
    Dim position As Long
    Dim lowestAt As Long
    Dim assigned As Long

    Set live = New Collection
    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then live.Add componentValue
    Next componentValue
    Do While live.Count > 0
        Set lowest = live.Item(1)
        lowestAt = 1
        For position = 2 To live.Count
            If live.Item(position).ToastSlot < lowest.ToastSlot Then
                Set lowest = live.Item(position)
                lowestAt = position
            End If
        Next position
        assigned = assigned + 1
        If lowest.ToastSlot <> assigned Then lowest.ReslotToast assigned
        live.Remove lowestAt
    Loop
End Sub

Friend Function HasAppWork() As Boolean
    Dim opValue As ReDimUI
    Dim jobValue As ReDimUI
    Dim componentValue As ReDimUI

    For Each opValue In mAsyncOps
        If opValue.OpIsRunning Then
            HasAppWork = True
            Exit Function
        End If
    Next opValue
    For Each jobValue In mJobs
        If jobValue.JobIsRunning Then
            HasAppWork = True
            Exit Function
        End If
    Next jobValue
    For Each componentValue In mComponents
        If componentValue.NeedsAnimation Then
            HasAppWork = True
            Exit Function
        End If
        If componentValue.IsPendingToast Then
            HasAppWork = True
            Exit Function
        End If
        If componentValue.IsSlideDragging Then
            HasAppWork = True
            Exit Function
        End If
        If componentValue.NeedsDragWatch(mSheet) Then
            HasAppWork = True
            Exit Function
        End If
    Next componentValue
End Function

' ---------------------------------------------------------------
' Async ops and jobs
' ---------------------------------------------------------------

Public Function Async(ByVal opId As String) As ReDimUI
    Dim opValue As ReDimUI

    RequireRole ROLE_APP, "Async"
    RequireIdentifier opId, True, "Async", "opId"
    Set opValue = FindKeyed(mAsyncOps, opId)
    If opValue Is Nothing Then
        Set opValue = New ReDimUI
        opValue.ConfigureAsyncOp mAppId, opId
        mAsyncOps.Add opValue, opId
    End If
    Set Async = opValue
End Function

Public Function Job(ByVal jobId As String) As ReDimUI
    Dim jobValue As ReDimUI

    RequireRole ROLE_APP, "Job"
    RequireIdentifier jobId, True, "Job", "jobId"
    Set jobValue = FindKeyed(mJobs, jobId)
    If jobValue Is Nothing Then
        Set jobValue = New ReDimUI
        jobValue.ConfigureJob mAppId, jobId
        mJobs.Add jobValue, jobId
    End If
    Set Job = jobValue
End Function

Private Function FindKeyed(ByVal source As Collection, ByVal key As String) As ReDimUI
    On Error Resume Next
    Set FindKeyed = source.Item(key)
    On Error GoTo 0
End Function

Public Sub CancelAsync(ByVal opId As String)
    Dim opValue As ReDimUI

    RequireRole ROLE_APP, "CancelAsync"
    Set opValue = FindKeyed(mAsyncOps, opId)
    If opValue Is Nothing Then
        RaiseError ERROR_NOT_FOUND, "CancelAsync", _
            "No async op is named '" & opId & "'."
    Else
        opValue.RequestOpCancel
    End If
End Sub

Public Sub CancelJob(ByVal jobId As String)
    Dim jobValue As ReDimUI

    RequireRole ROLE_APP, "CancelJob"
    Set jobValue = FindKeyed(mJobs, jobId)
    If jobValue Is Nothing Then
        RaiseError ERROR_NOT_FOUND, "CancelJob", _
            "No job is named '" & jobId & "'."
    Else
        jobValue.RequestJobCancel
    End If
End Sub

Public Function AsyncError(ByVal opId As String) As String
    Dim opValue As ReDimUI

    RequireRole ROLE_APP, "AsyncError"
    Set opValue = FindKeyed(mAsyncOps, opId)
    If Not opValue Is Nothing Then AsyncError = opValue.OpLastError
End Function

Friend Function IsOpRunning(ByVal opId As String) As Boolean
    Dim opValue As ReDimUI

    Set opValue = FindKeyed(mAsyncOps, opId)
    If Not opValue Is Nothing Then IsOpRunning = opValue.OpIsRunning
End Function

' =====================================================================
' Component configuration and fluent surface
' =====================================================================

Friend Sub ConfigureComponent( _
    ByVal ownerAppId As String, _
    ByVal componentId As String, _
    ByVal componentKind As Long _
)
    mRole = ROLE_COMPONENT
    mOwnerAppId = ownerAppId
    mComponentId = componentId
    mKind = componentKind
    mVisible = True
    mEnabled = True
    mFillColor = -1
    mTextColor = -1
    mFontSize = 0
    mBusyText = "Working..."
    Select Case componentKind
        Case KIND_BUTTON
            mWidth = 96: mHeight = 32
        Case KIND_LABEL
            mWidth = 200: mHeight = 20
        Case KIND_CARD
            mWidth = 200: mHeight = 110
        Case KIND_SPINNER
            mWidth = 24: mHeight = 24
            mVisible = False
        Case KIND_PROGRESS
            mWidth = 200: mHeight = 12
        Case KIND_TOGGLE
            mWidth = 44: mHeight = 22
        Case KIND_INPUT
            mWidth = 140: mHeight = 20
        Case KIND_COMBO
            mWidth = 150: mHeight = 20
        Case KIND_SELECT
            mWidth = 130: mHeight = 24
        Case KIND_TICKBOX
            mWidth = 130: mHeight = 18
        Case KIND_RADIO
            mWidth = 140: mHeight = 60
        Case KIND_STEPPER
            mWidth = 120: mHeight = 24
            mSliderMin = 0: mSliderMax = 100: mSliderStep = 1
        Case KIND_SLIDEBAR
            mWidth = 160: mHeight = 18
            mSliderMin = 0: mSliderMax = 100: mSliderStep = 1
        Case KIND_TOAST
            mWidth = 240: mHeight = 40
        Case Else
            mWidth = 120: mHeight = 24
    End Select
    mLeft = 10: mTop = 10
End Sub

Public Property Get ComponentId() As String
    RequireRole ROLE_COMPONENT, "ComponentId"
    ComponentId = mComponentId
End Property

Friend Property Get ComponentKind() As Long
    ComponentKind = mKind
End Property

Public Property Get OwnerApp() As ReDimUI
    RequireRole ROLE_COMPONENT, "OwnerApp"
    Set OwnerApp = ReDimUI.App(mOwnerAppId)
End Property

Public Function At(ByVal anchor As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "At"
    mAnchor = anchor
    mHasRect = False
    Touch
    Set At = Me
End Function

Public Function AtRect( _
    ByVal leftPoints As Double, _
    ByVal topPoints As Double, _
    ByVal widthPoints As Double, _
    ByVal heightPoints As Double _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "AtRect"
    mAnchor = vbNullString
    mRelMode = REL_NONE
    mLeft = leftPoints
    mTop = topPoints
    mWidth = widthPoints
    mHeight = heightPoints
    mHasRect = True
    Touch
    Set AtRect = Me
End Function

' Relative placement: position under another component, left edges aligned,
' keeping this component's own size.
Public Function Below( _
    ByVal otherComponentId As String, _
    Optional ByVal gapPoints As Double = 8 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "Below"
    mRelMode = REL_BELOW
    mRelTargetId = otherComponentId
    mRelGap = gapPoints
    mAnchor = vbNullString
    Touch
    Set Below = Me
End Function

' Relative placement: position to the right of another component, top edges
' aligned, keeping this component's own size.
Public Function RightOf( _
    ByVal otherComponentId As String, _
    Optional ByVal gapPoints As Double = 8 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "RightOf"
    mRelMode = REL_RIGHT_OF
    mRelTargetId = otherComponentId
    mRelGap = gapPoints
    mAnchor = vbNullString
    Touch
    Set RightOf = Me
End Function

Public Function Sized( _
    ByVal widthPoints As Double, _
    ByVal heightPoints As Double _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "Sized"
    mWidth = widthPoints
    mHeight = heightPoints
    Touch
    Set Sized = Me
End Function

Public Function Text(ByVal value As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Text"
    mText = value
    Touch
    Set Text = Me
End Function

Public Function BusyText(ByVal value As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BusyText"
    mBusyText = value
    Set BusyText = Me
End Function

Public Function FontSize(ByVal points As Double) As ReDimUI
    RequireRole ROLE_COMPONENT, "FontSize"
    mFontSize = points
    Touch
    Set FontSize = Me
End Function

Public Function Bold(Optional ByVal value As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Bold"
    mFontBold = value
    Touch
    Set Bold = Me
End Function

Public Function Fill(ByVal color As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "Fill"
    mFillColor = color
    mStyleVariant = VARIANT_NEUTRAL
    Touch
    Set Fill = Me
End Function

Public Function TextColor(ByVal color As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "TextColor"
    mTextColor = color
    Touch
    Set TextColor = Me
End Function

Public Function Primary() As ReDimUI
    Set Primary = SetVariant(VARIANT_PRIMARY)
End Function

Public Function Secondary() As ReDimUI
    Set Secondary = SetVariant(VARIANT_SECONDARY)
End Function

Public Function Success() As ReDimUI
    Set Success = SetVariant(VARIANT_SUCCESS)
End Function

Public Function Danger() As ReDimUI
    Set Danger = SetVariant(VARIANT_DANGER)
End Function

Private Function SetVariant(ByVal value As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "Variant"
    mStyleVariant = value
    Touch
    Set SetVariant = Me
End Function

Public Function Visible(ByVal value As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Visible"
    mVisible = value
    Touch
    Set Visible = Me
End Function

Public Function Enabled(ByVal value As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Enabled"
    mEnabled = value
    Touch
    Set Enabled = Me
End Function

Public Function Value(ByVal newValue As Double) As ReDimUI
    RequireRole ROLE_COMPONENT, "Value"
    mValue = newValue
    Touch
    Set Value = Me
End Function

Public Property Get CurrentValue() As Double
    RequireRole ROLE_COMPONENT, "CurrentValue"
    CurrentValue = mValue
End Property

Public Property Get CurrentText() As String
    RequireRole ROLE_COMPONENT, "CurrentText"
    CurrentText = ResolveText(OwnerApp)
End Property

Public Property Get IsEnabled() As Boolean
    RequireRole ROLE_COMPONENT, "IsEnabled"
    IsEnabled = mEnabled
End Property

Public Property Get IsBusy() As Boolean
    RequireRole ROLE_COMPONENT, "IsBusy"
    IsBusy = mBusy
End Property

' Registers a zero-argument workbook procedure run on click.
Public Function OnClick(ByVal procName As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnClick"
    mClickProc = procName
    Touch
    Set OnClick = Me
End Function

' Registers a zero-argument workbook procedure run through the async engine:
' the button disables itself and shows busy text while the work runs.
Public Function OnClickAsync(ByVal procName As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnClickAsync"
    mAsyncClickProc = procName
    Touch
    Set OnClickAsync = Me
End Function

' One-declaration navigation link: clicking this component navigates to
' the named window after any OnClick handler runs.
Public Function NavigatesTo(ByVal appId As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "NavigatesTo"
    mNavTargetId = appId
    Touch
    Set NavigatesTo = Me
End Function

Public Function BindText( _
    ByVal stateKey As String, _
    Optional ByVal formatTemplate As String _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindText"
    mBindTextKey = stateKey
    mBindTextFormat = formatTemplate
    Touch
    Set BindText = Me
End Function

Public Function BindValue(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindValue"
    mBindValueKey = stateKey
    Touch
    Set BindValue = Me
End Function

' invert:=True shows the component while the key is False. The disable-
' while-busy pattern binds Enabled to a running flag with invert.
Public Function BindVisible( _
    ByVal stateKey As String, _
    Optional ByVal invert As Boolean = False _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindVisible"
    mBindVisibleKey = stateKey
    mBindVisibleInvert = invert
    Touch
    Set BindVisible = Me
End Function

Public Function BindEnabled( _
    ByVal stateKey As String, _
    Optional ByVal invert As Boolean = False _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindEnabled"
    mBindEnabledKey = stateKey
    mBindEnabledInvert = invert
    Touch
    Set BindEnabled = Me
End Function

Public Function WritesTo(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "WritesTo"
    mWritesToKey = stateKey
    Set WritesTo = Me
End Function

Public Function Checked(ByVal value As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Checked"
    mChecked = value
    Touch
    Set Checked = Me
End Function

Public Property Get IsChecked() As Boolean
    RequireRole ROLE_COMPONENT, "IsChecked"
    IsChecked = mChecked
End Property

Public Function Items(ParamArray listItems() As Variant) As ReDimUI
    Dim position As Long

    RequireRole ROLE_COMPONENT, "Items"
    Set mItems = New Collection
    For position = LBound(listItems) To UBound(listItems)
        mItems.Add CStr(listItems(position))
    Next position
    ClampItemSelection
    Touch
    Set Items = Me
End Function

' Replaces the item list from a 1D array, a Collection, or a Range (one
' item per non-empty cell), so pickers can be fed straight from data.
Public Function ItemsFrom(ByVal source As Variant) As ReDimUI
    Dim entry As Variant
    Dim cell As Range
    Dim position As Long

    RequireRole ROLE_COMPONENT, "ItemsFrom"
    Set mItems = New Collection
    If TypeName(source) = "Range" Then
        For Each cell In source.Cells
            If LenB(CStr(cell.Value)) > 0 Then mItems.Add CStr(cell.Value)
        Next cell
    ElseIf TypeName(source) = "Collection" Then
        For Each entry In source
            mItems.Add CStr(entry)
        Next entry
    ElseIf IsArray(source) Then
        For position = LBound(source) To UBound(source)
            mItems.Add CStr(source(position))
        Next position
    Else
        RaiseError ERROR_INVALID_ARGUMENT, "ItemsFrom", _
            "ItemsFrom accepts an array, a Collection, or a Range."
    End If
    ClampItemSelection
    Touch
    Set ItemsFrom = Me
End Function

' Appends, or inserts at a 1-based position. The selected item stays
' selected: an insert at or before it shifts the index along with it.
Public Function AddItem( _
    ByVal itemText As String, _
    Optional ByVal atPosition As Long = 0 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "AddItem"
    If mItems Is Nothing Then Set mItems = New Collection
    If atPosition >= 1 And atPosition <= mItems.Count Then
        mItems.Add itemText, Before:=atPosition
        If CLng(mValue) >= atPosition Then mValue = mValue + 1
    Else
        mItems.Add itemText
    End If
    Touch
    Set AddItem = Me
End Function

' Removes by 1-based position, or by first matching text. Removing an item
' before the selection shifts the index; removing the selected item clears
' the selection to the placeholder rather than guessing a neighbor.
Public Function RemoveItem(ByVal indexOrText As Variant) As ReDimUI
    Dim position As Long
    Dim probe As Long

    RequireRole ROLE_COMPONENT, "RemoveItem"
    If mItems Is Nothing Then
        Set RemoveItem = Me
        Exit Function
    End If
    If IsNumeric(indexOrText) Then
        position = CLng(indexOrText)
    Else
        For probe = 1 To mItems.Count
            If CStr(mItems.Item(probe)) = CStr(indexOrText) Then
                position = probe
                Exit For
            End If
        Next probe
    End If
    If position < 1 Or position > mItems.Count Then
        Set RemoveItem = Me
        Exit Function
    End If
    mItems.Remove position
    If CLng(mValue) = position Then
        mValue = 0
    ElseIf CLng(mValue) > position Then
        mValue = mValue - 1
    End If
    Touch
    Set RemoveItem = Me
End Function

Public Function ClearItems() As ReDimUI
    RequireRole ROLE_COMPONENT, "ClearItems"
    Set mItems = New Collection
    mValue = 0
    Touch
    Set ClearItems = Me
End Function

Public Property Get ItemCount() As Long
    RequireRole ROLE_COMPONENT, "ItemCount"
    If Not mItems Is Nothing Then ItemCount = mItems.Count
End Property

Public Property Get ItemTextAt(ByVal position As Long) As String
    RequireRole ROLE_COMPONENT, "ItemTextAt"
    ItemTextAt = ItemAt(position)
End Property

Private Sub ClampItemSelection()
    If mItems Is Nothing Then
        mValue = 0
    ElseIf CLng(mValue) > mItems.Count Then
        mValue = 0
    End If
End Sub

Public Function SliderRange( _
    ByVal minimum As Long, _
    ByVal maximum As Long, _
    Optional ByVal stepSize As Long = 1 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "SliderRange"
    If maximum <= minimum Then
        RaiseError ERROR_INVALID_ARGUMENT, "SliderRange", _
            "maximum must be greater than minimum."
    End If
    mSliderMin = minimum
    mSliderMax = maximum
    mSliderStep = stepSize
    Touch
    Set SliderRange = Me
End Function

' OnChange is OnClick under a truthful name for value controls: form control
' OnAction fires after the value changes.
Public Function OnChange(ByVal procName As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnChange"
    mClickProc = procName
    Touch
    Set OnChange = Me
End Function

' Reads or writes the backing cell of a TextInput without re-entering the
' SheetChange pipeline.
Public Property Get InputValue() As String
    RequireRole ROLE_COMPONENT, "InputValue"
    If IsFloatField() Then
        InputValue = mText
    Else
        InputValue = CStr(InputCell(OwnerApp.Sheet).Value)
    End If
End Property

Public Property Let InputValue(ByVal newValue As String)
    Dim priorEvents As Boolean

    RequireRole ROLE_COMPONENT, "InputValue"
    If IsFloatField() Then
        mText = newValue
        Touch
    Else
        priorEvents = Application.EnableEvents
        Application.EnableEvents = False
        InputCell(OwnerApp.Sheet).Value = newValue
        Application.EnableEvents = priorEvents
    End If
    If LenB(mWritesToKey) > 0 Then OwnerApp.SetState mWritesToKey, newValue
End Property

Friend Function InputCell(ByVal hostSheet As Worksheet) As Range
    If LenB(mAnchor) = 0 Then
        RaiseError ERROR_INVALID_OPERATION, "TextInput", _
            "TextInput '" & mComponentId & "' needs At with a cell anchor."
    End If
    Set InputCell = hostSheet.Range(mAnchor).Cells(1, 1)
End Function

Friend Sub SyncInputFromSheet(ByVal ownerApp As ReDimUI, ByVal changed As Range)
    Dim cell As Range
    Dim newValue As String
    Dim exactAt As Long

    If mKind <> KIND_INPUT And mKind <> KIND_COMBO Then Exit Sub
    If LenB(mAnchor) = 0 Then Exit Sub
    On Error Resume Next
    Set cell = InputCell(ownerApp.Sheet)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub
    If Application.Intersect(cell, changed) Is Nothing Then Exit Sub
    newValue = CStr(cell.Value)
    If LenB(mWritesToKey) > 0 Then ownerApp.SetState mWritesToKey, newValue
    If LenB(mClickProc) > 0 Then
        On Error Resume Next
        Application.Run mClickProc
        On Error GoTo 0
    End If
    ' Combo auto-suggest at commit time, the moment Excel hands control
    ' back: an exact match takes the item and closes; partial matches drop
    ' the filtered list; no match stays closed as free text.
    If mKind = KIND_COMBO Then
        exactAt = ExactItemMatch(newValue)
        If exactAt > 0 Then
            mValue = exactAt
            mSelectOpen = False
        ElseIf LenB(newValue) > 0 _
            And BuildComboMatches(newValue).Count > 0 Then
            mValue = 0
            mSelectOpen = True
        Else
            mValue = 0
            mSelectOpen = False
        End If
        Touch
    End If
End Sub

Friend Sub ConfigureToast( _
    ByVal message As String, _
    ByVal ttlMs As Long, _
    ByVal slot As Long _
)
    Dim hostApp As ReDimUI
    Dim boundLeft As Double
    Dim boundTop As Double
    Dim boundRight As Double
    Dim viewWidth As Double
    Dim viewHeight As Double

    mText = message
    mExpiresAtMs = GetTickCount64() + ttlMs
    mValue = slot
    Set hostApp = OwnerApp
    ' Tray origin, in preference order: an explicit ToastTray anchor (used
    ' exactly as pinned), or a rail just outside the content's right edge,
    ' clamped into the visible viewport. A toast is a notification first:
    ' when the rail and the viewport disagree, one of them hides the
    ' message, so visibility wins in exactly that case and the rail wins in
    ' every other.
    Dim originLeft As Double
    Dim originTop As Double
    Dim viewLeft As Double
    Dim viewTop As Double
    If Not hostApp.ToastTrayOrigin(originLeft, originTop) Then
        If hostApp.ContentBounds(boundLeft, boundTop, boundRight) Then
            originLeft = boundRight + 12
            originTop = boundTop
        Else
            hostApp.ResolveViewport boundLeft, boundTop, viewWidth, viewHeight
            originLeft = boundLeft + viewWidth - mWidth - 12
            originTop = boundTop + 8
        End If
        hostApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
        If originLeft + mWidth + 12 > viewLeft + viewWidth Then
            originLeft = viewLeft + viewWidth - mWidth - 12
        End If
        If originLeft < viewLeft + 12 Then originLeft = viewLeft + 12
        If originTop + mHeight + 8 > viewTop + viewHeight Then
            originTop = viewTop + viewHeight - mHeight - 8
        End If
        If originTop < viewTop + 8 Then originTop = viewTop + 8
    End If
    mLeft = originLeft
    mTop = originTop + (slot - 1) * (mHeight + 6)
    ' Entrance animation: the shape spawns this far below its slot and the
    ' pump eases it up into place.
    mSlideRemaining = 14
    mHasRect = True
    mVisible = True
End Sub

Friend Property Get ToastSlot() As Long
    If mKind = KIND_TOAST Then ToastSlot = CLng(mValue)
End Property

' Moves this toast to a new slot, easing from wherever it currently sits
' (including mid-entrance) to the new position without a visual jump.
Friend Sub ReslotToast(ByVal newSlot As Long)
    Dim spacing As Double
    Dim originTop As Double
    Dim visualTop As Double

    If mKind <> KIND_TOAST Then Exit Sub
    spacing = mHeight + 6
    originTop = mTop - (CLng(mValue) - 1) * spacing
    visualTop = mTop + mSlideRemaining
    mValue = newSlot
    mTop = originTop + (newSlot - 1) * spacing
    mSlideRemaining = visualTop - mTop
    If mSlideRemaining < 0 Then mSlideRemaining = 0
    ReDimUI.RequestPump
End Sub

Friend Function IsExpiredToast(ByVal nowMs As LongLong) As Boolean
    If mKind <> KIND_TOAST Then Exit Function
    IsExpiredToast = (nowMs >= mExpiresAtMs)
End Function

Friend Property Get IsPendingToast() As Boolean
    IsPendingToast = (mKind = KIND_TOAST)
End Property

' Modal chrome (the reserved mdl_ ids) is viewport-centered and transient,
' so it never counts toward content bounds: a toast fired while a modal is
' up must not shift the tray.
Friend Property Get CountsForBounds() As Boolean
    CountsForBounds = (mKind <> KIND_TOAST And mKind <> KIND_OVERLAY _
        And mVisible And Left$(mComponentId, 4) <> "mdl_")
End Property

Friend Property Get BoundsLeft() As Double
    BoundsLeft = mLeft
End Property

Friend Property Get BoundsTop() As Double
    BoundsTop = mTop
End Property

Friend Property Get BoundsRight() As Double
    BoundsRight = mLeft + mWidth
End Property

Friend Sub MarkCloseModalOnClick()
    mCloseModalOnClick = True
End Sub

Public Sub Remove()
    Dim hostApp As ReDimUI

    RequireRole ROLE_COMPONENT, "Remove"
    Set hostApp = OwnerApp
    DeleteShapes hostApp.Sheet
    hostApp.ForgetComponent mComponentId
End Sub

Friend Sub ForgetComponent(ByVal componentId As String)
    Dim position As Long

    RequireRole ROLE_APP, "ForgetComponent"
    On Error Resume Next
    mComponents.Remove componentId
    On Error GoTo 0
    For position = mComponentOrder.Count To 1 Step -1
        If mComponentOrder.Item(position) = componentId Then
            mComponentOrder.Remove position
        End If
    Next position
End Sub

Friend Function IsBoundTo(ByVal stateKey As String) As Boolean
    IsBoundTo = (mBindTextKey = stateKey) Or (mBindValueKey = stateKey) _
        Or (mBindVisibleKey = stateKey) Or (mBindEnabledKey = stateKey)
End Function

Friend Function CanClick() As Boolean
    If Not mVisible Or Not mEnabled Or mBusy Then Exit Function
    If GetTickCount64() - mLastClickMs < CLICK_DEBOUNCE_MS Then Exit Function
    CanClick = True
End Function

Friend Sub MarkClicked()
    mLastClickMs = GetTickCount64()
End Sub

' Locals here must not reuse member names: VBA identifiers are
' case-insensitive, so a local named ownerApp would shadow the OwnerApp
' property and Set ownerApp = OwnerApp would self-assign Nothing.
Friend Sub RunClickBehavior()
    Dim hostApp As ReDimUI

    Set hostApp = OwnerApp
    Select Case mKind
        Case KIND_TOGGLE
            mChecked = Not mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
            Touch
        Case KIND_INPUT
            If IsFloatField() Then
                FocusField hostApp
                Exit Sub
            End If
        Case KIND_COMBO
            If Left$(ReDimUI.SenderPart, 3) = "opt" Then
                PickComboItem hostApp, CLng(Mid$(ReDimUI.SenderPart, 4))
            ElseIf IsFloatField() Then
                If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
                    mSelectOpen = Not mSelectOpen
                    Touch
                Else
                    FocusField hostApp
                End If
                Exit Sub
            Else
                mSelectOpen = Not mSelectOpen
                Touch
                Exit Sub
            End If
        Case KIND_SELECT
            If Left$(ReDimUI.SenderPart, 3) = "opt" Then
                mValue = CLng(Mid$(ReDimUI.SenderPart, 4))
                mSelectOpen = False
                If LenB(mWritesToKey) > 0 Then
                    hostApp.SetState mWritesToKey, ItemAt(CLng(mValue))
                End If
                Touch
            Else
                mSelectOpen = Not mSelectOpen
                Touch
                Exit Sub
            End If
        Case KIND_TICKBOX
            mChecked = Not mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
            Touch
        Case KIND_RADIO
            Dim rowIndex As Long
            If LenB(ReDimUI.SenderPart) = 0 Then
                rowIndex = 1
            ElseIf IsNumeric(Mid$(ReDimUI.SenderPart, 2)) Then
                rowIndex = CLng(Mid$(ReDimUI.SenderPart, 2))
            End If
            If rowIndex < 1 Then Exit Sub
            If rowIndex = CLng(mValue) Then Exit Sub
            mValue = rowIndex
            If LenB(mWritesToKey) > 0 Then
                hostApp.SetState mWritesToKey, ItemAt(rowIndex)
            End If
            Touch
        Case KIND_STEPPER
            Dim stepped As Double
            If ReDimUI.SenderPart = "minus" Then
                stepped = mValue - mSliderStep
            ElseIf ReDimUI.SenderPart = "plus" Then
                stepped = mValue + mSliderStep
            Else
                Exit Sub
            End If
            If stepped < mSliderMin Then stepped = mSliderMin
            If stepped > mSliderMax Then stepped = mSliderMax
            If stepped = mValue Then Exit Sub
            mValue = stepped
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mValue
            Touch
        Case KIND_SLIDEBAR
            ' Drags are handled by the pump's press watch; the release
            ' click arrives here and is swallowed by the suppress window.
            ' A tap with the pump disarmed still sets via this fallback.
            If Not mSlideTracking Then
                If GetTickCount64() >= mSlideSuppressUntilMs Then
                    HandleSlideClick hostApp
                End If
            End If
            Exit Sub
        Case KIND_TOAST
            mExpiresAtMs = 1
            ReDimUI.RequestPump
            Exit Sub
    End Select
    If LenB(mAsyncClickProc) > 0 Then
        hostApp.StartAsyncClick Me, mAsyncClickProc
        Exit Sub
    End If
    If LenB(mClickProc) > 0 Then Application.Run mClickProc
    If mCloseModalOnClick Then hostApp.CloseModal
    If LenB(mNavTargetId) > 0 Then ReDimUI.Navigate mNavTargetId
End Sub

Private Function ItemAt(ByVal position As Long) As String
    If mItems Is Nothing Then Exit Function
    If position < 1 Or position > mItems.Count Then Exit Function
    ItemAt = CStr(mItems.Item(position))
End Function

' The OnClickAsync path: a per-button op that disables the button, runs the
' procedure through the pump, and restores the button on completion.
Friend Sub StartAsyncClick(ByVal source As ReDimUI, ByVal procName As String)
    Dim opValue As ReDimUI

    RequireRole ROLE_APP, "StartAsyncClick"
    Set opValue = Async("click_" & source.ComponentId)
    If opValue.OpIsRunning Then Exit Sub
    opValue.ConfigureClickOp procName, source.ComponentId
    opValue.Start
End Sub

Friend Sub SetBusy(ByVal value As Boolean)
    mBusy = value
    Touch
End Sub

Private Sub Touch()
    Dim hostApp As ReDimUI

    If LenB(mOwnerAppId) = 0 Then Exit Sub
    If Not ReDimUI.HasApp(mOwnerAppId) Then Exit Sub
    Set hostApp = ReDimUI.App(mOwnerAppId)
    hostApp.NotifyComponentTouched mComponentId
End Sub

' =====================================================================
' Component rendering
' =====================================================================

Friend Property Get ShapeName() As String
    ShapeName = SHAPE_PREFIX & mOwnerAppId & "_" & mComponentId
End Property

Friend Sub Apply( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal ownerApp As ReDimUI _
)
    Dim target As Shape

    RequireRole ROLE_COMPONENT, "Apply"
    ResolveBindings ownerApp
    EnsureGeometryResolved hostSheet
    Set target = EnsureShape(hostSheet)
    ApplyGeometry target
    ApplyStyle target, themeValue
    ApplyText target, themeValue
    ApplyBehavior target, hostSheet
    ApplyKindExtras target, hostSheet, themeValue
    mAppliedValid = True
End Sub

Private Sub ResolveBindings(ByVal ownerApp As ReDimUI)
    Dim bound As Variant

    If LenB(mBindTextKey) > 0 Then
        If ownerApp.HasState(mBindTextKey) Then
            mText = FormatBoundText(ownerApp.State(mBindTextKey))
        End If
    End If
    If LenB(mBindValueKey) > 0 Then
        If ownerApp.HasState(mBindValueKey) Then
            bound = ownerApp.State(mBindValueKey)
            If IsNumeric(bound) Then mValue = CDbl(bound)
        End If
    End If
    If LenB(mBindVisibleKey) > 0 Then
        If ownerApp.HasState(mBindVisibleKey) Then
            mVisible = CBool(ownerApp.State(mBindVisibleKey)) Xor mBindVisibleInvert
        End If
    End If
    If LenB(mBindEnabledKey) > 0 Then
        If ownerApp.HasState(mBindEnabledKey) Then
            mEnabled = CBool(ownerApp.State(mBindEnabledKey)) Xor mBindEnabledInvert
        End If
    End If
End Sub

Private Function FormatBoundText(ByVal boundValue As Variant) As String
    If LenB(mBindTextFormat) = 0 Then
        FormatBoundText = CStr(boundValue)
    Else
        FormatBoundText = Replace(mBindTextFormat, "{0}", CStr(boundValue))
    End If
End Function

Private Sub ResolveGeometry(ByVal hostSheet As Worksheet)
    Dim anchorRange As Range
    Dim relTarget As ReDimUI

    If mRelMode <> REL_NONE Then
        If Not OwnerApp.TryGetComponent(mRelTargetId, relTarget) Then
            RaiseError ERROR_NOT_FOUND, "Below", _
                "Relative anchor '" & mRelTargetId & "' is not a component of this app."
        End If
        relTarget.EnsureGeometryResolved hostSheet
        If mRelMode = REL_BELOW Then
            mLeft = relTarget.BoundsLeft
            mTop = relTarget.BoundsTop + relTarget.BoundsHeight + mRelGap
        Else
            mLeft = relTarget.BoundsRight + mRelGap
            mTop = relTarget.BoundsTop
        End If
        Exit Sub
    End If
    If LenB(mAnchor) = 0 Then Exit Sub
    On Error GoTo BadAnchor
    Set anchorRange = hostSheet.Range(mAnchor)
    On Error GoTo 0
    mLeft = anchorRange.Left
    mTop = anchorRange.Top
    mWidth = anchorRange.Width
    mHeight = anchorRange.Height
    Exit Sub

BadAnchor:
    RaiseError ERROR_INVALID_ARGUMENT, "At", _
        "'" & mAnchor & "' is not a valid range on sheet '" & hostSheet.Name & "'."
End Sub

' Relative chains resolve the anchor's geometry first. The flag turns a
' circular chain into a clear error instead of infinite recursion.
Friend Sub EnsureGeometryResolved(ByVal hostSheet As Worksheet)
    If mGeometryResolving Then
        RaiseError ERROR_INVALID_OPERATION, "Below", _
            "Relative placement of '" & mComponentId & "' is circular."
    End If
    mGeometryResolving = True
    On Error GoTo Unwind
    ResolveGeometry hostSheet
    mGeometryResolving = False
    Exit Sub

Unwind:
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim capturedDescription As String
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    mGeometryResolving = False
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Sub

Friend Property Get BoundsHeight() As Double
    BoundsHeight = mHeight
End Property

Private Function EnsureShape(ByVal hostSheet As Worksheet) As Shape
    Dim target As Shape

    On Error Resume Next
    Set target = hostSheet.Shapes(ShapeName)
    On Error GoTo 0
    If Not target Is Nothing Then
        ' An adopted shape from an earlier framework version can carry the
        ' wrong geometry type; rebuild it rather than styling the wrong form.
        If mKind = KIND_SPINNER Then
            If target.AutoShapeType <> msoShapeBlockArc Then
                target.Delete
                Set target = Nothing
            End If
        End If
    End If
    If target Is Nothing Then
        Set target = CreateShapeForKind(hostSheet)
        target.Name = ShapeName
        mAppliedValid = False
    End If
    Set EnsureShape = target
End Function

Private Function CreateShapeForKind(ByVal hostSheet As Worksheet) As Shape
    Dim target As Shape

    Select Case mKind
        Case KIND_BUTTON
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.25
        Case KIND_LABEL
            Set target = hostSheet.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, mLeft, mTop, mWidth, mHeight)
        Case KIND_CARD
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.08
        Case KIND_SPINNER
            ' A block arc is a filled ring segment concentric with its
            ' bounding box, so rotation stays visually circular. A plain
            ' arc's path center is off the box center and wobbles.
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeBlockArc, mLeft, mTop, mWidth, mHeight)
            On Error Resume Next
            target.Adjustments(1) = 0
            target.Adjustments(2) = 290
            target.Adjustments(3) = 0.34
            On Error GoTo 0
        Case KIND_TOGGLE
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.5
        Case KIND_TOAST
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.2
        Case KIND_SELECT
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.15
        Case KIND_TICKBOX
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, 14, 14)
            target.Adjustments(1) = 0.2
        Case KIND_RADIO
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeOval, mLeft, mTop, 14, 14)
        Case KIND_STEPPER
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.15
        Case KIND_SLIDEBAR
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, 6)
            target.Adjustments(1) = 0.5
        Case KIND_INPUT, KIND_COMBO
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            target.Adjustments(1) = 0.15
        Case Else
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeRectangle, mLeft, mTop, mWidth, mHeight)
    End Select
    target.Placement = xlFreeFloating
    Set CreateShapeForKind = target
End Function

Private Function EnsurePartShape( _
    ByVal hostSheet As Worksheet, _
    ByVal partName As String, _
    ByVal shapeType As MsoAutoShapeType _
) As Shape
    Dim target As Shape
    Dim fullName As String

    fullName = ShapeName & PART_SEPARATOR & partName
    On Error Resume Next
    Set target = hostSheet.Shapes(fullName)
    On Error GoTo 0
    If target Is Nothing Then
        Set target = hostSheet.Shapes.AddShape(shapeType, mLeft, mTop, 8, 8)
        target.Name = fullName
        target.Placement = xlFreeFloating
        target.Line.Visible = msoFalse
    End If
    Set EnsurePartShape = target
End Function

Private Function IsCompositeDrawnKind() As Boolean
    IsCompositeDrawnKind = (mKind = KIND_TICKBOX Or mKind = KIND_RADIO _
        Or mKind = KIND_STEPPER Or mKind = KIND_SLIDEBAR)
End Function

' Geometry diffs against the shape's actual position rather than a cache, so
' a shape moved by hand (Ctrl+drag) snaps back to its declared rectangle on
' the next render. Composite drawn kinds own their geometry in extras: the
' main shape is one piece of several, not the declared rectangle.
Private Sub ApplyGeometry(ByVal target As Shape)
    If IsCompositeDrawnKind() Then Exit Sub
    If Abs(target.Left - mLeft) > 0.01 Then target.Left = mLeft
    If Abs(target.Top - mTop) > 0.01 Then target.Top = mTop
    If Abs(target.Width - mWidth) > 0.01 Then target.Width = mWidth
    If Abs(target.Height - mHeight) > 0.01 Then target.Height = mHeight
    If Not mAppliedValid Or mAppliedVisible <> mVisible Then
        target.Visible = BoolToMso(mVisible)
        mAppliedVisible = mVisible
    End If
End Sub

Private Sub ApplyStyle(ByVal target As Shape, ByVal themeValue As ReDimUI)
    Dim resolvedFill As Long

    If IsCompositeDrawnKind() Then Exit Sub
    resolvedFill = ResolveFill(themeValue)
    Select Case mKind
        Case KIND_LABEL
            If Not mAppliedValid Then
                target.Fill.Visible = msoFalse
                target.Line.Visible = msoFalse
            End If
        Case KIND_SPINNER
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = ResolveSpinnerColor(themeValue)
                target.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
        Case KIND_CARD, KIND_TOAST
            If Not mAppliedValid Or mAppliedFill <> ResolveSurfaceFill(themeValue) Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                target.Line.Visible = msoTrue
                target.Line.ForeColor.RGB = themeValue.BorderColor
                target.Line.Weight = 1
                mAppliedFill = ResolveSurfaceFill(themeValue)
            End If
        Case KIND_PROGRESS
            If Not mAppliedValid Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = themeValue.MutedColor
                target.Line.Visible = msoFalse
            End If
        Case KIND_TOGGLE
            If Not mAppliedValid Or mAppliedFill <> ResolveToggleFill(themeValue) Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = ResolveToggleFill(themeValue)
                target.Line.Visible = msoFalse
                mAppliedFill = ResolveToggleFill(themeValue)
            End If
        Case KIND_INPUT, KIND_COMBO
            ' Shapes adopted from builds older than 0.9.1 are square;
            ' coerce them to the rounded profile the family shares.
            If Not mAppliedValid Then
                If target.AutoShapeType <> msoShapeRoundedRectangle Then
                    target.AutoShapeType = msoShapeRoundedRectangle
                    target.Adjustments(1) = 0.15
                End If
            End If
            If IsFloatField() Then
                ' Float fields are their own editor: surface fill so the
                ' face takes clicks, accent border while focused.
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                target.Line.Visible = msoTrue
                If ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
                    target.Line.ForeColor.RGB = themeValue.PrimaryColor
                    target.Line.Weight = 1.5
                Else
                    target.Line.ForeColor.RGB = themeValue.BorderColor
                    target.Line.Weight = 1
                End If
            ElseIf Not mAppliedValid Then
                target.Fill.Visible = msoFalse
                target.Line.Visible = msoTrue
                target.Line.ForeColor.RGB = themeValue.BorderColor
                target.Line.Weight = 1
            End If
        Case KIND_SELECT
            If Not mAppliedValid Or mAppliedFill <> ResolveSurfaceFill(themeValue) Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                target.Line.Visible = msoTrue
                target.Line.ForeColor.RGB = themeValue.BorderColor
                target.Line.Weight = 1
                mAppliedFill = ResolveSurfaceFill(themeValue)
            End If
        Case KIND_OVERLAY
            If Not mAppliedValid Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = RGB(0, 0, 0)
                target.Fill.Transparency = 0.45
                target.Line.Visible = msoFalse
            End If
        Case Else
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                target.Fill.Visible = msoTrue
                target.Fill.ForeColor.RGB = resolvedFill
                ' Surface-filled buttons need an outline or they vanish
                ' against cards and the canvas.
                If mStyleVariant = VARIANT_SECONDARY And mEnabled And Not mBusy Then
                    target.Line.Visible = msoTrue
                    target.Line.ForeColor.RGB = themeValue.BorderColor
                    target.Line.Weight = 1
                Else
                    target.Line.Visible = msoFalse
                End If
                mAppliedFill = resolvedFill
            End If
    End Select
End Sub

Private Function ResolveSurfaceFill(ByVal themeValue As ReDimUI) As Long
    If mFillColor >= 0 Then
        ResolveSurfaceFill = mFillColor
    Else
        ResolveSurfaceFill = themeValue.SurfaceColor
    End If
End Function

Private Function ResolveToggleFill(ByVal themeValue As ReDimUI) As Long
    If mChecked Then
        ResolveToggleFill = themeValue.PrimaryColor
    Else
        ResolveToggleFill = themeValue.MutedColor
    End If
End Function

Private Function ResolveFill(ByVal themeValue As ReDimUI) As Long
    If Not mEnabled Or mBusy Then
        ResolveFill = themeValue.MutedColor
        Exit Function
    End If
    If mFillColor >= 0 Then
        ResolveFill = mFillColor
        Exit Function
    End If
    Select Case mStyleVariant
        Case VARIANT_PRIMARY
            ResolveFill = themeValue.PrimaryColor
        Case VARIANT_SECONDARY
            ResolveFill = themeValue.SurfaceColor
        Case VARIANT_SUCCESS
            ResolveFill = themeValue.SuccessColor
        Case VARIANT_DANGER
            ResolveFill = themeValue.DangerColor
        Case Else
            If mKind = KIND_CARD Then
                ResolveFill = themeValue.SurfaceColor
            Else
                ResolveFill = themeValue.MutedColor
            End If
    End Select
End Function

Private Function ResolveTextColor(ByVal themeValue As ReDimUI) As Long
    If Not mEnabled Or mBusy Then
        ResolveTextColor = themeValue.OnMutedColor
        Exit Function
    End If
    If mTextColor >= 0 Then
        ResolveTextColor = mTextColor
        Exit Function
    End If
    Select Case mKind
        Case KIND_LABEL, KIND_CARD, KIND_SELECT, KIND_TOAST, _
            KIND_INPUT, KIND_COMBO
            ' Surface-filled kinds read in surface ink regardless of variant;
            ' the variant fall-through below is for filled buttons only.
            ' Float fields especially: their implicit variant is primary,
            ' which would put the primary's white ink on the surface fill
            ' and render typed text and the caret invisible.
            ResolveTextColor = themeValue.OnSurfaceColor
        Case Else
            Select Case mStyleVariant
                Case VARIANT_SECONDARY, VARIANT_NEUTRAL
                    ResolveTextColor = themeValue.OnSurfaceColor
                Case Else
                    ResolveTextColor = themeValue.OnPrimaryColor
            End Select
    End Select
End Function

Friend Function ResolveText(ByVal ownerApp As ReDimUI) As String
    If mBusy And mKind = KIND_BUTTON And LenB(mBusyText) > 0 Then
        ResolveText = mBusyText
        Exit Function
    End If
    If mKind = KIND_SELECT Then
        ResolveText = ItemAt(CLng(mValue))
        If LenB(ResolveText) = 0 Then ResolveText = mText
        Exit Function
    End If
    ' A focused float field shows the blinking insertion bar in-text.
    If IsFloatField() Then
        If ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
            ResolveText = mText & IIf(mCaretOn, "|", " ")
            Exit Function
        End If
    End If
    ResolveText = mText
End Function

Private Function ResolveSpinnerColor(ByVal themeValue As ReDimUI) As Long
    If mFillColor >= 0 Then
        ResolveSpinnerColor = mFillColor
    Else
        ResolveSpinnerColor = themeValue.PrimaryColor
    End If
End Function

Private Sub ApplyText(ByVal target As Shape, ByVal themeValue As ReDimUI)
    Dim resolved As String
    Dim resolvedColor As Long
    Dim resolvedSize As Double
    Dim textChanged As Boolean

    If mKind = KIND_SPINNER Or mKind = KIND_TOGGLE Or mKind = KIND_OVERLAY _
        Or IsCompositeDrawnKind() Then Exit Sub
    If (mKind = KIND_INPUT Or mKind = KIND_COMBO) And Not IsFloatField() Then
        Exit Sub
    End If
    resolved = ResolveText(Nothing)
    resolvedColor = ResolveTextColor(themeValue)
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    If Not mAppliedValid Then
        target.TextFrame2.MarginLeft = 6
        target.TextFrame2.MarginRight = 6
        target.TextFrame2.MarginTop = 3
        target.TextFrame2.MarginBottom = 3
        target.TextFrame2.VerticalAnchor = msoAnchorMiddle
        target.TextFrame2.WordWrap = msoTrue
        If mKind = KIND_BUTTON Then
            target.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Else
            target.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End If
        If mKind = KIND_CARD Then
            target.TextFrame2.VerticalAnchor = msoAnchorTop
        End If
    End If
    textChanged = (Not mAppliedValid) Or (mAppliedText <> resolved)
    If textChanged Then
        target.TextFrame2.TextRange.Text = resolved
        mAppliedText = resolved
    End If
    ' Replacing shape text can drop character formatting, so the font is
    ' reapplied on any text change, not only on style changes.
    If textChanged Or mAppliedTextColor <> resolvedColor _
        Or mAppliedFontSize <> resolvedSize Or mAppliedFontBold <> mFontBold Then
        With target.TextFrame2.TextRange.Font
            .Name = themeValue.FontName
            .Size = resolvedSize
            .Bold = BoolToMso(mFontBold Or mKind = KIND_BUTTON)
            .Fill.ForeColor.RGB = resolvedColor
        End With
        mAppliedTextColor = resolvedColor
        mAppliedFontSize = resolvedSize
        mAppliedFontBold = mFontBold
    End If
End Sub

' Every framework shape routes through the dispatcher, including kinds with
' no click behavior. A shape with a macro assigned runs it instead of being
' selected, so a plain click can never drag a modal card or any other piece
' of chrome out of position. Ctrl+click still selects for design work.
Private Sub ApplyBehavior(ByVal target As Shape, ByVal hostSheet As Worksheet)
    Dim desiredAction As String

    desiredAction = "'" & hostSheet.Parent.Name & "'!RdxDispatch"
    If Not mAppliedValid Or mAppliedOnAction <> desiredAction Then
        target.OnAction = desiredAction
        mAppliedOnAction = desiredAction
    End If
End Sub

' Composite and control-backed widgets keep their extra pieces in sync here:
' the toggle knob, the progress fill, form control values, and the input
' cell's one-time formatting. Part geometry is synced unconditionally; the
' shapes are tiny and correctness beats caching for moving pieces.
Private Sub ApplyKindExtras( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Select Case mKind
        Case KIND_TOGGLE
            ApplyToggleKnob hostSheet, themeValue
            mAppliedChecked = mChecked
        Case KIND_PROGRESS
            ApplyProgressFill hostSheet, themeValue
            mAppliedValue = mValue
        Case KIND_INPUT
            If Not IsFloatField() Then ApplyInputCell hostSheet, themeValue
        Case KIND_COMBO
            If Not IsFloatField() Then ApplyInputCell hostSheet, themeValue
            ApplyComboParts hostSheet, themeValue
        Case KIND_SELECT
            ApplySelectParts hostSheet, themeValue
        Case KIND_TICKBOX
            ApplyTickBox target, hostSheet, themeValue
        Case KIND_RADIO
            ApplyRadioGroup target, hostSheet, themeValue
        Case KIND_STEPPER
            ApplyStepper target, hostSheet, themeValue
        Case KIND_SLIDEBAR
            ApplySlideBar target, hostSheet, themeValue
        Case KIND_TOAST
            If mSlideRemaining > 0 Then target.Top = mTop + mSlideRemaining
    End Select
End Sub

Private Sub StyleGlyphText( _
    ByVal target As Shape, _
    ByVal glyph As String, _
    ByVal inkColor As Long, _
    ByVal glyphSize As Double, _
    ByVal themeValue As ReDimUI _
)
    With target.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = glyph
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = glyphSize
            .Bold = msoTrue
            .Fill.ForeColor.RGB = inkColor
        End With
    End With
End Sub

Private Sub StyleCaptionPart( _
    ByVal caption As Shape, _
    ByVal captionText As String, _
    ByVal themeValue As ReDimUI _
)
    Dim resolvedSize As Double

    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    caption.Fill.Visible = msoFalse
    caption.Line.Visible = msoFalse
    With caption.TextFrame2
        .MarginLeft = 2
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = captionText
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = resolvedSize
            .Bold = BoolToMso(mFontBold)
            .Fill.ForeColor.RGB = ResolveDrawnInk(themeValue)
        End With
    End With
End Sub

Private Function ResolveDrawnInk(ByVal themeValue As ReDimUI) As Long
    If mEnabled Then
        ResolveDrawnInk = themeValue.OnSurfaceColor
    Else
        ResolveDrawnInk = themeValue.OnMutedColor
    End If
End Function

' Box at the left edge, check glyph inside it when checked, caption filling
' the remainder. Both pieces click-toggle through the dispatcher.
Private Sub ApplyTickBox( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim caption As Shape
    Dim boxSize As Double

    boxSize = mHeight - 2
    If boxSize > 16 Then boxSize = 16
    If boxSize < 10 Then boxSize = 10
    target.Left = mLeft
    target.Top = mTop + (mHeight - boxSize) / 2
    target.Width = boxSize
    target.Height = boxSize
    target.Visible = BoolToMso(mVisible)
    target.Fill.Visible = msoTrue
    If mChecked And mEnabled Then
        target.Fill.ForeColor.RGB = ResolveFill(themeValue)
        target.Line.Visible = msoFalse
    Else
        target.Fill.ForeColor.RGB = themeValue.SurfaceColor
        target.Line.Visible = msoTrue
        target.Line.ForeColor.RGB = themeValue.BorderColor
        target.Line.Weight = 1.25
    End If
    If mChecked Then
        StyleGlyphText target, ChrW(10003), _
            IIf(mEnabled, themeValue.OnPrimaryColor, themeValue.OnMutedColor), _
            boxSize * 0.62, themeValue
    Else
        StyleGlyphText target, vbNullString, themeValue.OnSurfaceColor, _
            boxSize * 0.62, themeValue
    End If
    Set caption = EnsurePartShape(hostSheet, "lbl", msoShapeRectangle)
    caption.Left = mLeft + boxSize + 6
    caption.Top = mTop - 1
    caption.Width = mWidth - boxSize - 6
    caption.Height = mHeight + 2
    StyleCaptionPart caption, mText, themeValue
    caption.Visible = BoolToMso(mVisible)
    caption.OnAction = mAppliedOnAction
    mAppliedChecked = mChecked
End Sub

' One circle-and-caption row per item. The selected row's circle fills with
' the accent and carries an inner dot part.
Private Sub ApplyRadioGroup( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim rowCount As Long
    Dim rowHeight As Double
    Dim rowTop As Double
    Dim position As Long
    Dim ring As Shape
    Dim caption As Shape
    Dim dot As Shape
    Dim circleSize As Double
    Dim selected As Boolean

    If mItems Is Nothing Then Exit Sub
    rowCount = mItems.Count
    If rowCount = 0 Then Exit Sub
    rowHeight = mHeight / rowCount
    circleSize = rowHeight - 6
    If circleSize > 14 Then circleSize = 14
    If circleSize < 10 Then circleSize = 10
    For position = 1 To rowCount
        rowTop = mTop + (position - 1) * rowHeight
        selected = (position = CLng(mValue))
        If position = 1 Then
            Set ring = target
        Else
            Set ring = EnsurePartShape(hostSheet, "c" & position, msoShapeOval)
        End If
        ring.Left = mLeft
        ring.Top = rowTop + (rowHeight - circleSize) / 2
        ring.Width = circleSize
        ring.Height = circleSize
        ring.Fill.Visible = msoTrue
        ring.Fill.ForeColor.RGB = themeValue.SurfaceColor
        ring.Line.Visible = msoTrue
        If selected And mEnabled Then
            ring.Line.ForeColor.RGB = ResolveFill(themeValue)
            ring.Line.Weight = 1.5
        Else
            ring.Line.ForeColor.RGB = themeValue.BorderColor
            ring.Line.Weight = 1.25
        End If
        ring.Visible = BoolToMso(mVisible)
        ring.OnAction = mAppliedOnAction
        Set dot = EnsurePartShape(hostSheet, "d" & position, msoShapeOval)
        dot.Left = ring.Left + circleSize * 0.3
        dot.Top = ring.Top + circleSize * 0.3
        dot.Width = circleSize * 0.4
        dot.Height = circleSize * 0.4
        dot.Fill.Visible = msoTrue
        If mEnabled Then
            dot.Fill.ForeColor.RGB = ResolveFill(themeValue)
        Else
            dot.Fill.ForeColor.RGB = themeValue.OnMutedColor
        End If
        dot.Line.Visible = msoFalse
        dot.Visible = BoolToMso(mVisible And selected)
        dot.OnAction = mAppliedOnAction
        Set caption = EnsurePartShape(hostSheet, "t" & position, msoShapeRectangle)
        caption.Left = mLeft + circleSize + 8
        caption.Top = rowTop
        caption.Width = mWidth - circleSize - 8
        caption.Height = rowHeight
        StyleCaptionPart caption, ItemAt(position), themeValue
        caption.Visible = BoolToMso(mVisible)
        caption.OnAction = mAppliedOnAction
    Next position
    ' Rows beyond the current item count are leftovers from a longer list.
    For position = rowCount + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "c" & position).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "d" & position).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & position).Delete
        On Error GoTo 0
    Next position
    mAppliedOptionCount = rowCount
    mAppliedValue = mValue
End Sub

' Minus and plus squares around a bordered value face.
Private Sub ApplyStepper( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim minusButton As Shape
    Dim plusButton As Shape
    Dim buttonSize As Double

    buttonSize = mHeight
    target.Left = mLeft + buttonSize + 4
    target.Top = mTop
    target.Width = mWidth - 2 * (buttonSize + 4)
    target.Height = mHeight
    target.Visible = BoolToMso(mVisible)
    target.Fill.Visible = msoTrue
    target.Fill.ForeColor.RGB = themeValue.SurfaceColor
    target.Line.Visible = msoTrue
    target.Line.ForeColor.RGB = themeValue.BorderColor
    target.Line.Weight = 1
    StyleGlyphText target, CStr(CLng(mValue)), ResolveDrawnInk(themeValue), _
        IIf(mFontSize > 0, mFontSize, themeValue.BaseFontSize), themeValue

    Set minusButton = EnsurePartShape(hostSheet, "minus", msoShapeRoundedRectangle)
    minusButton.Left = mLeft
    minusButton.Top = mTop
    minusButton.Width = buttonSize
    minusButton.Height = buttonSize
    minusButton.Fill.Visible = msoTrue
    minusButton.Fill.ForeColor.RGB = themeValue.MutedColor
    minusButton.Line.Visible = msoFalse
    StyleGlyphText minusButton, "-", ResolveDrawnInk(themeValue), _
        themeValue.BaseFontSize + 2, themeValue
    minusButton.Visible = BoolToMso(mVisible)
    minusButton.OnAction = mAppliedOnAction

    Set plusButton = EnsurePartShape(hostSheet, "plus", msoShapeRoundedRectangle)
    plusButton.Left = mLeft + mWidth - buttonSize
    plusButton.Top = mTop
    plusButton.Width = buttonSize
    plusButton.Height = buttonSize
    plusButton.Fill.Visible = msoTrue
    plusButton.Fill.ForeColor.RGB = themeValue.MutedColor
    plusButton.Line.Visible = msoFalse
    StyleGlyphText plusButton, "+", ResolveDrawnInk(themeValue), _
        themeValue.BaseFontSize + 2, themeValue
    plusButton.Visible = BoolToMso(mVisible)
    plusButton.OnAction = mAppliedOnAction
    mAppliedValue = mValue
End Sub

Private Sub ApplyToggleKnob( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim knob As Shape
    Dim knobSize As Double

    Set knob = EnsurePartShape(hostSheet, "knob", msoShapeOval)
    knobSize = mHeight - 4
    If knobSize < 4 Then knobSize = 4
    knob.Width = knobSize
    knob.Height = knobSize
    knob.Top = mTop + 2
    If mChecked Then
        knob.Left = mLeft + mWidth - knobSize - 2
    Else
        knob.Left = mLeft + 2
    End If
    knob.Fill.Visible = msoTrue
    knob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    knob.Visible = BoolToMso(mVisible)
    knob.OnAction = mAppliedOnAction
End Sub

Private Sub ApplyProgressFill( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim fillPart As Shape
    Dim clamped As Double

    Set fillPart = EnsurePartShape(hostSheet, "fill", msoShapeRectangle)
    clamped = mValue
    If clamped < 0 Then clamped = 0
    If clamped > 100 Then clamped = 100
    fillPart.Left = mLeft
    fillPart.Top = mTop
    fillPart.Height = mHeight
    If clamped <= 0 Then
        fillPart.Visible = msoFalse
    Else
        fillPart.Width = mWidth * clamped / 100
        fillPart.Visible = BoolToMso(mVisible)
    End If
    fillPart.Fill.Visible = msoTrue
    fillPart.Fill.ForeColor.RGB = ResolveFill(themeValue)
End Sub

Private Sub ApplySelectParts( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim caret As Shape

    Set caret = EnsurePartShape(hostSheet, "caret", msoShapeIsoscelesTriangle)
    caret.Width = 8
    caret.Height = 5
    caret.Left = mLeft + mWidth - 14
    caret.Top = mTop + (mHeight - 5) / 2
    caret.Rotation = 180
    caret.Fill.Visible = msoTrue
    caret.Fill.ForeColor.RGB = themeValue.OnMutedColor
    caret.Visible = BoolToMso(mVisible)
    caret.OnAction = mAppliedOnAction
    SyncSelectOptions hostSheet, themeValue
End Sub

Private Sub SyncSelectOptions( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim optionShape As Shape
    Dim position As Long
    Dim wanted As Long
    Dim resolvedSize As Double

    If mSelectOpen And mVisible And Not mItems Is Nothing Then
        wanted = mItems.Count
    End If
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    For position = 1 To wanted
        Set optionShape = EnsurePartShape( _
            hostSheet, "opt" & position, msoShapeRectangle)
        optionShape.Left = mLeft
        optionShape.Top = mTop + mHeight + (position - 1) * mHeight
        optionShape.Width = mWidth
        optionShape.Height = mHeight
        optionShape.Fill.Visible = msoTrue
        If position = CLng(mValue) Then
            optionShape.Fill.ForeColor.RGB = themeValue.MutedColor
        Else
            optionShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        End If
        optionShape.Line.Visible = msoTrue
        optionShape.Line.ForeColor.RGB = themeValue.BorderColor
        optionShape.Line.Weight = 1
        With optionShape.TextFrame2
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 2
            .MarginBottom = 2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .TextRange.Text = ItemAt(position)
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            With .TextRange.Font
                .Name = themeValue.FontName
                .Size = resolvedSize
                .Bold = msoFalse
                .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
            End With
        End With
        optionShape.Visible = msoTrue
        optionShape.OnAction = mAppliedOnAction
        optionShape.ZOrder msoBringToFront
    Next position
    For position = wanted + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & position).Delete
        On Error GoTo 0
    Next position
    mAppliedOptionCount = wanted
End Sub

' Slim centered track, accent fill up to the value, and a round thumb.
Private Sub ApplySlideBar( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim fillPart As Shape
    Dim thumb As Shape
    Dim trackTop As Double
    Dim fraction As Double
    Dim thumbSize As Double

    fraction = SlideFraction()
    trackTop = mTop + (mHeight - 6) / 2
    target.Left = mLeft
    target.Top = trackTop
    target.Width = mWidth
    target.Height = 6
    target.Visible = BoolToMso(mVisible)
    target.Fill.Visible = msoTrue
    target.Fill.ForeColor.RGB = themeValue.MutedColor
    target.Line.Visible = msoFalse

    Set fillPart = EnsurePartShape(hostSheet, "fill", msoShapeRoundedRectangle)
    fillPart.Left = mLeft
    fillPart.Top = trackTop
    fillPart.Height = 6
    If fraction <= 0 Then
        fillPart.Visible = msoFalse
    Else
        fillPart.Width = mWidth * fraction
        fillPart.Visible = BoolToMso(mVisible)
    End If
    fillPart.Fill.Visible = msoTrue
    If mEnabled Then
        fillPart.Fill.ForeColor.RGB = ResolveFill(themeValue)
    Else
        fillPart.Fill.ForeColor.RGB = themeValue.OnMutedColor
    End If
    fillPart.Line.Visible = msoFalse
    fillPart.OnAction = mAppliedOnAction

    thumbSize = mHeight
    Set thumb = EnsurePartShape(hostSheet, "thumb", msoShapeOval)
    thumb.Width = thumbSize
    thumb.Height = thumbSize
    thumb.Left = mLeft + mWidth * fraction - thumbSize / 2
    If thumb.Left < mLeft - thumbSize / 2 Then thumb.Left = mLeft - thumbSize / 2
    thumb.Top = mTop
    thumb.Fill.Visible = msoTrue
    ' The engaged thumb carries the accent so tracking mode is visible.
    If mSlideTracking Then
        thumb.Fill.ForeColor.RGB = ResolveFill(themeValue)
    Else
        thumb.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    thumb.Line.Visible = msoTrue
    thumb.Line.ForeColor.RGB = themeValue.BorderColor
    thumb.Line.Weight = 1.25
    thumb.Visible = BoolToMso(mVisible)
    thumb.OnAction = mAppliedOnAction
    mAppliedValue = mValue
End Sub

Private Function SlideFraction() As Double
    If mSliderMax <= mSliderMin Then Exit Function
    SlideFraction = (mValue - mSliderMin) / (mSliderMax - mSliderMin)
    If SlideFraction < 0 Then SlideFraction = 0
    If SlideFraction > 1 Then SlideFraction = 1
End Function

' Applies a track fraction to the value with SliderRange snapping. The drag
' loop calls it live; tests call it directly as a deterministic seam.
Friend Sub SlideToFraction( _
    ByVal hostApp As ReDimUI, _
    ByVal fraction As Double, _
    ByVal fireChange As Boolean _
)
    Dim rawValue As Double
    Dim snapped As Double

    If fraction < 0 Then fraction = 0
    If fraction > 1 Then fraction = 1
    rawValue = mSliderMin + fraction * (mSliderMax - mSliderMin)
    If mSliderStep > 0 Then
        snapped = mSliderMin + _
            CLng((rawValue - mSliderMin) / mSliderStep) * mSliderStep
    Else
        snapped = rawValue
    End If
    If snapped < mSliderMin Then snapped = mSliderMin
    If snapped > mSliderMax Then snapped = mSliderMax
    If snapped <> mValue Then
        mValue = snapped
        If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mValue
        Touch
    End If
    If fireChange Then
        If LenB(mClickProc) > 0 Then
            On Error Resume Next
            Application.Run mClickProc
            On Error GoTo 0
        End If
    End If
End Sub

' Click-to-set: one cursor read at dispatch time, snapped and applied.
' Deliberately loop-free; nothing in ReDim may block the pump, and shape
' OnAction fires on mouse up anyway, so there is no held button to track.
Friend Sub HandleSlideClick(ByVal hostApp As ReDimUI)
    Dim fraction As Double

    fraction = CursorTrackFraction()
    If fraction < 0 Then Exit Sub
    SlideToFraction hostApp, fraction, True
End Sub

' True press-drag, watched by the pump: OnAction only fires at mouse up,
' so the frames themselves detect the left-button press edge, hit-test the
' cursor against this slider's rectangle, and run the drag session - value
' tracks while the button is held, finalizes on release. One key-state
' poll and at most one cursor read per frame; nothing blocks.
Friend Sub TickSlideDrag(ByVal hostApp As ReDimUI, ByVal buttonDown As Boolean)
    Dim fraction As Double

    If mKind <> KIND_SLIDEBAR Then Exit Sub
    If mSlideTracking Then
        If Not mVisible Or Not mEnabled Then
            EndSlideDrag hostApp
        ElseIf buttonDown Then
            fraction = CursorTrackFraction()
            If fraction >= 0 Then ContinueSlideDrag hostApp, fraction
        Else
            EndSlideDrag hostApp
        End If
    ElseIf buttonDown And Not mSlidePrevButtonDown Then
        If mVisible And mEnabled Then
            If CursorOverTrack() Then
                fraction = CursorTrackFraction()
                If fraction >= 0 Then BeginSlideDrag hostApp, fraction
            End If
        End If
    End If
    mSlidePrevButtonDown = buttonDown
End Sub

Friend Sub BeginSlideDrag(ByVal hostApp As ReDimUI, ByVal fraction As Double)
    mSlideTracking = True
    mSlideValueAtEngage = mValue
    SlideToFraction hostApp, fraction, False
    Touch
End Sub

Friend Sub ContinueSlideDrag(ByVal hostApp As ReDimUI, ByVal fraction As Double)
    If Not mSlideTracking Then Exit Sub
    SlideToFraction hostApp, fraction, False
End Sub

' Release: fire OnChange once if the session moved the value, and swallow
' the OnAction click Excel is about to deliver for the same release.
Friend Sub EndSlideDrag(ByVal hostApp As ReDimUI)
    If Not mSlideTracking Then Exit Sub
    mSlideTracking = False
    mSlideSuppressUntilMs = GetTickCount64() + 400
    Touch
    If mValue <> mSlideValueAtEngage Then
        If LenB(mClickProc) > 0 Then
            On Error Resume Next
            Application.Run mClickProc
            On Error GoTo 0
        End If
    End If
End Sub

Friend Property Get IsSlideDragging() As Boolean
    IsSlideDragging = mSlideTracking
End Property

' A slider on the active sheet keeps the pump armed so the press edge is
' never missed. Idle cost is one key-state poll per frame.
Friend Function NeedsDragWatch(ByVal hostSheet As Worksheet) As Boolean
    If mKind <> KIND_SLIDEBAR Then Exit Function
    If Not mVisible Or Not mEnabled Then Exit Function
    On Error Resume Next
    NeedsDragWatch = (hostSheet Is ActiveSheet)
    On Error GoTo 0
End Function

Private Function CursorOverTrack() As Boolean
    Dim cursor As RdxScreenPoint
    Dim pointsX As Double
    Dim pointsY As Double

    If GetCursorPos(cursor) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(cursor.X)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(cursor.Y)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    If pointsX < mLeft - 8 Or pointsX > mLeft + mWidth + 8 Then Exit Function
    If pointsY < mTop - 4 Or pointsY > mTop + mHeight + 4 Then Exit Function
    CursorOverTrack = True
End Function

Private Function CursorTrackFraction() As Double
    Dim cursor As RdxScreenPoint
    Dim cursorPoints As Double

    CursorTrackFraction = -1
    If GetCursorPos(cursor) = 0 Then Exit Function
    cursorPoints = ScreenXToSheetPoints(cursor.X)
    If cursorPoints = UNMAPPABLE_POINTS Then Exit Function
    If mWidth <= 0 Then Exit Function
    CursorTrackFraction = (cursorPoints - mLeft) / mWidth
    If CursorTrackFraction < 0 Then CursorTrackFraction = 0
    If CursorTrackFraction > 1 Then CursorTrackFraction = 1
End Function

' Maps a screen pixel X to absolute sheet points. PointsToScreenPixelsX is
' misnamed on current Excel builds: its input is document PIXELS (measured
' slope exactly 1 at any zoom, while scrolling shifts the origin by the
' true pixels-per-point). The conversion is therefore chosen empirically:
' if the measured slope matches DPI times zoom it is a real points input;
' if the slope is 1 the input is pixels and DPI times zoom converts. Under
' either contract (px - f(0)) / conversion yields absolute points, scroll
' already accounted for by f(0).
Friend Function ScreenXToSheetPoints(ByVal screenX As Long) As Double
    Dim originPx As Double
    Dim referencePx As Double
    Dim measuredSlope As Double
    Dim dpiSlope As Double
    Dim conversion As Double

    ScreenXToSheetPoints = UNMAPPABLE_POINTS
    On Error GoTo Unmappable
    originPx = ActiveWindow.PointsToScreenPixelsX(0)
    referencePx = ActiveWindow.PointsToScreenPixelsX(1000)
    measuredSlope = (referencePx - originPx) / 1000#
    dpiSlope = ScreenPixelsPerPoint()
    On Error GoTo 0
    If measuredSlope <= 0 Or dpiSlope <= 0 Then Exit Function
    If Abs(measuredSlope - dpiSlope) < 0.02 Then
        conversion = measuredSlope
    ElseIf Abs(measuredSlope - 1#) < 0.02 Then
        conversion = dpiSlope
    Else
        conversion = measuredSlope
    End If
    ScreenXToSheetPoints = (screenX - originPx) / conversion
    Exit Function

Unmappable:
    ScreenXToSheetPoints = UNMAPPABLE_POINTS
End Function

Friend Function ScreenYToSheetPoints(ByVal screenY As Long) As Double
    Dim originPx As Double
    Dim referencePx As Double
    Dim measuredSlope As Double
    Dim dpiSlope As Double
    Dim conversion As Double

    ScreenYToSheetPoints = UNMAPPABLE_POINTS
    On Error GoTo Unmappable
    originPx = ActiveWindow.PointsToScreenPixelsY(0)
    referencePx = ActiveWindow.PointsToScreenPixelsY(1000)
    measuredSlope = (referencePx - originPx) / 1000#
    dpiSlope = ScreenPixelsPerPointY()
    On Error GoTo 0
    If measuredSlope <= 0 Or dpiSlope <= 0 Then Exit Function
    If Abs(measuredSlope - dpiSlope) < 0.02 Then
        conversion = measuredSlope
    ElseIf Abs(measuredSlope - 1#) < 0.02 Then
        conversion = dpiSlope
    Else
        conversion = measuredSlope
    End If
    ScreenYToSheetPoints = (screenY - originPx) / conversion
    Exit Function

Unmappable:
    ScreenYToSheetPoints = UNMAPPABLE_POINTS
End Function

Private Function ScreenPixelsPerPointY() As Double
    Static dotsPerInch As Double
    Dim deviceContext As LongPtr

    If dotsPerInch = 0 Then
        deviceContext = GetDC(0)
        If deviceContext <> 0 Then
            dotsPerInch = GetDeviceCaps(deviceContext, LOGPIXELSY)
            ReleaseDC 0, deviceContext
        End If
        If dotsPerInch = 0 Then dotsPerInch = 96
    End If
    On Error Resume Next
    ScreenPixelsPerPointY = dotsPerInch / 72# * ActiveWindow.Zoom / 100#
    On Error GoTo 0
    If ScreenPixelsPerPointY = 0 Then ScreenPixelsPerPointY = dotsPerInch / 72#
End Function

Private Function ScreenPixelsPerPoint() As Double
    Static dotsPerInch As Double
    Dim deviceContext As LongPtr

    If dotsPerInch = 0 Then
        deviceContext = GetDC(0)
        If deviceContext <> 0 Then
            dotsPerInch = GetDeviceCaps(deviceContext, LOGPIXELSX)
            ReleaseDC 0, deviceContext
        End If
        If dotsPerInch = 0 Then dotsPerInch = 96
    End If
    On Error Resume Next
    ScreenPixelsPerPoint = dotsPerInch / 72# * ActiveWindow.Zoom / 100#
    On Error GoTo 0
    If ScreenPixelsPerPoint = 0 Then ScreenPixelsPerPoint = dotsPerInch / 72#
End Function

Private Sub ApplyComboParts( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim caret As Shape

    Set caret = EnsurePartShape(hostSheet, "caret", msoShapeIsoscelesTriangle)
    caret.Width = 8
    caret.Height = 5
    caret.Left = mLeft + mWidth - 12
    caret.Top = mTop + (mHeight - 5) / 2
    caret.Rotation = 180
    caret.Fill.Visible = msoTrue
    caret.Fill.ForeColor.RGB = themeValue.OnMutedColor
    caret.Visible = BoolToMso(mVisible)
    caret.OnAction = mAppliedOnAction
    SyncComboOptions hostSheet, themeValue
End Sub

' The open list shows items filtered by the typed text: contains-match,
' case-insensitive, everything when the text is empty. Option rows map
' back to item indexes through mComboMatches.
Private Sub SyncComboOptions( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim optionShape As Shape
    Dim position As Long
    Dim wanted As Long
    Dim itemIndex As Long
    Dim resolvedSize As Double

    Set mComboMatches = BuildComboMatches(FieldText(hostSheet))
    If mSelectOpen And mVisible Then wanted = mComboMatches.Count
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    For position = 1 To wanted
        itemIndex = CLng(mComboMatches.Item(position))
        Set optionShape = EnsurePartShape( _
            hostSheet, "opt" & position, msoShapeRectangle)
        optionShape.Left = mLeft
        optionShape.Top = mTop + mHeight + (position - 1) * (mHeight + 2)
        optionShape.Width = mWidth
        optionShape.Height = mHeight + 2
        optionShape.Fill.Visible = msoTrue
        If itemIndex = CLng(mValue) Then
            optionShape.Fill.ForeColor.RGB = themeValue.MutedColor
        Else
            optionShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        End If
        optionShape.Line.Visible = msoTrue
        optionShape.Line.ForeColor.RGB = themeValue.BorderColor
        optionShape.Line.Weight = 1
        With optionShape.TextFrame2
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 2
            .MarginBottom = 2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .TextRange.Text = ItemAt(itemIndex)
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            With .TextRange.Font
                .Name = themeValue.FontName
                .Size = resolvedSize
                .Bold = msoFalse
                .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
            End With
        End With
        optionShape.Visible = msoTrue
        optionShape.OnAction = mAppliedOnAction
        optionShape.ZOrder msoBringToFront
    Next position
    For position = wanted + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & position).Delete
        On Error GoTo 0
    Next position
    mAppliedOptionCount = wanted
End Sub

' A field without a cell anchor is a float field: its text lives in the
' component buffer, rendered on the shape itself, edited through the
' keyboard focus layer. Anchoring to a cell with At keeps the original
' cell-backed behavior.
Friend Function IsFloatField() As Boolean
    IsFloatField = ((mKind = KIND_INPUT Or mKind = KIND_COMBO) _
        And LenB(mAnchor) = 0)
End Function

Private Function FieldText(ByVal hostSheet As Worksheet) As String
    Dim cell As Range

    If IsFloatField() Then
        FieldText = mText
        Exit Function
    End If
    On Error Resume Next
    Set cell = InputCell(hostSheet)
    If Not cell Is Nothing Then FieldText = CStr(cell.Value)
    On Error GoTo 0
End Function

Friend Sub FocusField(ByVal hostApp As ReDimUI)
    If Not IsFloatField() Then Exit Sub
    mFieldTextAtFocus = mText
    mCaretOn = True
    mCaretAccumMs = 0
    ReDimUI.SetKeyboardFocus hostApp, Me
    If mKind = KIND_COMBO Then mSelectOpen = True
    Touch
End Sub

' Commit-blur keeps the text and fires events; revert-blur restores what
' was there when focus arrived and fires nothing.
Friend Sub BlurField(ByVal hostApp As ReDimUI, ByVal commitChanges As Boolean)
    Dim exactAt As Long

    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    ReDimUI.ClearKeyboardFocus
    mCaretOn = False
    If commitChanges Then
        If mKind = KIND_COMBO Then
            exactAt = ExactItemMatch(mText)
            If exactAt > 0 Then
                mValue = exactAt
                mText = ItemAt(exactAt)
            Else
                mValue = 0
            End If
        End If
        If mText <> mFieldTextAtFocus Then
            If LenB(mWritesToKey) > 0 Then
                hostApp.SetState mWritesToKey, mText
            End If
            If LenB(mClickProc) > 0 Then
                On Error Resume Next
                Application.Run mClickProc
                On Error GoTo 0
            End If
        End If
    Else
        mText = mFieldTextAtFocus
    End If
    mSelectOpen = False
    Touch
End Sub

Friend Sub HandleKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    If Not IsFloatField() Then Exit Sub
    Select Case keyText
        Case "{ENTER}", "{TAB}"
            ' Tab commits like Enter: the form-field convention for
            ' leaving a field with its value kept.
            BlurField hostApp, True
            Exit Sub
        Case "{ESC}"
            BlurField hostApp, False
            Exit Sub
        Case "{BS}"
            If Len(mText) > 0 Then mText = Left$(mText, Len(mText) - 1)
        Case Else
            mText = mText & keyText
    End Select
    mCaretOn = True
    mCaretAccumMs = 0
    ' Live filtering, the payoff cells could never give: the option list
    ' re-filters on every keystroke.
    If mKind = KIND_COMBO Then mSelectOpen = True
    Touch
End Sub

' Outside-press blur, watched by the same pump frames as slider drags: a
' press anywhere off the field commits and releases focus.
Friend Sub TickFieldWatch( _
    ByVal hostApp As ReDimUI, _
    ByVal buttonDown As Boolean, _
    ByVal pressEdge As Boolean _
)
    If Not IsFloatField() Then Exit Sub
    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    If Not pressEdge Then Exit Sub
    If CursorOverField() Then Exit Sub
    BlurField hostApp, True
End Sub

Private Function CursorOverField() As Boolean
    Dim cursor As RdxScreenPoint
    Dim pointsX As Double
    Dim pointsY As Double
    Dim listDepth As Double

    If GetCursorPos(cursor) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(cursor.X)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(cursor.Y)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    ' The open option list counts as part of the field.
    listDepth = 0
    If mSelectOpen And Not mComboMatches Is Nothing Then
        listDepth = mComboMatches.Count * (mHeight + 2)
    End If
    If pointsX < mLeft - 2 Or pointsX > mLeft + mWidth + 2 Then Exit Function
    If pointsY < mTop - 2 Or pointsY > mTop + mHeight + listDepth + 2 Then
        Exit Function
    End If
    CursorOverField = True
End Function

Private Function BuildComboMatches(ByVal filterText As String) As Collection
    Dim position As Long

    Set BuildComboMatches = New Collection
    If mItems Is Nothing Then Exit Function
    For position = 1 To mItems.Count
        If LenB(filterText) = 0 Then
            BuildComboMatches.Add position
        ElseIf InStr(1, CStr(mItems.Item(position)), filterText, _
            vbTextCompare) > 0 Then
            BuildComboMatches.Add position
        End If
    Next position
End Function

Private Function ExactItemMatch(ByVal candidate As String) As Long
    Dim position As Long

    If mItems Is Nothing Then Exit Function
    For position = 1 To mItems.Count
        If StrComp(CStr(mItems.Item(position)), candidate, _
            vbTextCompare) = 0 Then
            ExactItemMatch = position
            Exit Function
        End If
    Next position
End Function

' An option pick writes the cell without re-entering SheetChange, updates
' selection and state, and closes the list.
Friend Sub PickComboItem(ByVal hostApp As ReDimUI, ByVal matchRow As Long)
    Dim itemIndex As Long
    Dim priorEvents As Boolean
    Dim picked As String
    Dim cell As Range

    If mComboMatches Is Nothing Then Exit Sub
    If matchRow < 1 Or matchRow > mComboMatches.Count Then Exit Sub
    itemIndex = CLng(mComboMatches.Item(matchRow))
    picked = ItemAt(itemIndex)
    If IsFloatField() Then
        mText = picked
        If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
            ReDimUI.ClearKeyboardFocus
            mCaretOn = False
        End If
    Else
        priorEvents = Application.EnableEvents
        Application.EnableEvents = False
        On Error Resume Next
        Set cell = InputCell(hostApp.Sheet)
        If Not cell Is Nothing Then cell.Value = picked
        On Error GoTo 0
        Application.EnableEvents = priorEvents
    End If
    mValue = itemIndex
    mSelectOpen = False
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, picked
    Touch
End Sub

Private Sub ApplyInputCell( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim cell As Range

    If Not mAppliedValid Then
        Set cell = InputCell(hostSheet)
        cell.Font.Name = themeValue.FontName
        cell.Font.Size = themeValue.BaseFontSize
        cell.Font.Color = themeValue.OnSurfaceColor
        cell.Interior.Color = themeValue.SurfaceColor
        cell.HorizontalAlignment = xlLeft
        ' Unlocked from birth so ProtectSurface keeps inputs editable.
        cell.Locked = False
    End If
End Sub

Friend Sub UnlockInputCell(ByVal hostSheet As Worksheet)
    Dim cell As Range

    If mKind <> KIND_INPUT And mKind <> KIND_COMBO Then Exit Sub
    If LenB(mAnchor) = 0 Then Exit Sub
    On Error Resume Next
    Set cell = InputCell(hostSheet)
    If Not cell Is Nothing Then cell.Locked = False
    On Error GoTo 0
End Sub

Friend Sub DeleteShapes(ByVal hostSheet As Worksheet)
    Dim position As Long
    Dim optionCount As Long

    optionCount = mAppliedOptionCount
    If Not mItems Is Nothing Then
        If mItems.Count > optionCount Then optionCount = mItems.Count
    End If
    On Error Resume Next
    hostSheet.Shapes(ShapeName).Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "knob").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "fill").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "lbl").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "caret").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "minus").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "plus").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "thumb").Delete
    For position = 1 To optionCount
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & position).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "c" & position).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "d" & position).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & position).Delete
    Next position
    On Error GoTo 0
End Sub

Friend Function NeedsAnimation() As Boolean
    NeedsAnimation = (mKind = KIND_SPINNER And mVisible) _
        Or (mKind = KIND_TOAST And mSlideRemaining > 0) _
        Or (IsFloatField() And _
            ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId))
End Function

' Animation bypasses the diff cache: rotation and the toast slide are
' transient visual state, not declared props. Both are time-based so the
' frame rate changes smoothness, never speed: the spinner turns at 600
' degrees per second and the slide decays with a 50 ms half-life factor.
Friend Sub TickAnimation(ByVal hostSheet As Worksheet, ByVal elapsedMs As Long)
    If Not NeedsAnimation() Then Exit Sub
    On Error Resume Next
    Select Case mKind
        Case KIND_SPINNER
            hostSheet.Shapes(ShapeName).IncrementRotation elapsedMs * 0.6
        Case KIND_TOAST
            mSlideRemaining = mSlideRemaining * (0.55 ^ (elapsedMs / 50#))
            If mSlideRemaining < 0.75 Then mSlideRemaining = 0
            hostSheet.Shapes(ShapeName).Top = mTop + mSlideRemaining
        Case KIND_INPUT, KIND_COMBO
            ' Insertion bar blink for the focused float field.
            mCaretAccumMs = mCaretAccumMs + elapsedMs
            If mCaretAccumMs >= 500 Then
                mCaretAccumMs = 0
                mCaretOn = Not mCaretOn
                Touch
            End If
    End Select
    On Error GoTo 0
End Sub

' =====================================================================
' Async ops
' =====================================================================

Friend Sub ConfigureAsyncOp(ByVal ownerAppId As String, ByVal opId As String)
    mRole = ROLE_ASYNC_OP
    mOwnerAppId = ownerAppId
    mComponentId = opId
    Set mOpDisableIds = New Collection
End Sub

' Idempotent per-click configuration for the OnClickAsync path.
Friend Sub ConfigureClickOp(ByVal procName As String, ByVal componentId As String)
    RequireRole ROLE_ASYNC_OP, "ConfigureClickOp"
    mOpRunsProc = procName
    Set mOpTask = Nothing
    Set mOpDisableIds = New Collection
    mOpDisableIds.Add componentId
End Sub

Public Function RunsProc(ByVal procName As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "RunsProc"
    RequireNotRunning "RunsProc"
    mOpRunsProc = procName
    Set mOpTask = Nothing
    Set RunsProc = Me
End Function

Public Function RunsTask(ByVal taskValue As ROneCOne) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "RunsTask"
    RequireNotRunning "RunsTask"
    If taskValue Is Nothing Then
        RaiseError ERROR_INVALID_ARGUMENT, "RunsTask", "RunsTask requires a task."
    End If
    Set mOpTask = taskValue
    mOpRunsProc = vbNullString
    Set RunsTask = Me
End Function

Public Function WithCancellation() As ReDimUI
    RequireRole ROLE_ASYNC_OP, "WithCancellation"
    RequireNotRunning "WithCancellation"
    Set mOpCts = ROneCOne.CancellationTokenSource
    Set WithCancellation = Me
End Function

' The token to pass into task factories built outside the op, so CancelAsync
' reaches work created with RunsTask.
Public Property Get Token() As ROneCOne
    RequireRole ROLE_ASYNC_OP, "Token"
    If mOpCts Is Nothing Then
        RaiseError ERROR_INVALID_OPERATION, "Token", _
            "Call WithCancellation before reading Token."
    End If
    Set Token = mOpCts.Token
End Property

Public Function Disables(ParamArray componentIds() As Variant) As ReDimUI
    Dim position As Long

    RequireRole ROLE_ASYNC_OP, "Disables"
    RequireNotRunning "Disables"
    For position = LBound(componentIds) To UBound(componentIds)
        mOpDisableIds.Add CStr(componentIds(position))
    Next position
    Set Disables = Me
End Function

Public Function ShowsSpinner(ByVal spinnerId As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "ShowsSpinner"
    RequireNotRunning "ShowsSpinner"
    mOpSpinnerId = spinnerId
    Set ShowsSpinner = Me
End Function

Public Function OnDone(ByVal procName As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnDone"
    mOpDoneProc = procName
    Set OnDone = Me
End Function

Public Function OnFail(ByVal procName As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnFail"
    mOpFailProc = procName
    Set OnFail = Me
End Function

Public Function OnCancel(ByVal procName As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnCancel"
    mOpCancelProc = procName
    Set OnCancel = Me
End Function

Public Function TracksState(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "TracksState"
    mOpStatusKey = stateKey
    Set TracksState = Me
End Function

Public Sub Start()
    Dim hostApp As ReDimUI

    RequireRole ROLE_ASYNC_OP, "Start"
    RequireNotRunning "Start"
    Set hostApp = ReDimUI.App(mOwnerAppId)
    If LenB(mOpRunsProc) > 0 Then
        If mOpCts Is Nothing Then
            Set mOpTask = ROneCOne.Task.Run(ROneCOne.Action(mOpRunsProc))
        Else
            Set mOpTask = ROneCOne.Task.Run( _
                ROneCOne.Action(mOpRunsProc), mOpCts.Token)
        End If
    End If
    If mOpTask Is Nothing Then
        RaiseError ERROR_INVALID_OPERATION, "Start", _
            "Async op '" & mComponentId & "' needs RunsProc or RunsTask first."
    End If
    mOpLastError = vbNullString
    mOpRunning = True
    SetOpControlsBusy hostApp, True
    If LenB(mOpStatusKey) > 0 Then hostApp.SetState mOpStatusKey, "running"
    ReDimUI.RequestPump
End Sub

Friend Property Get OpIsRunning() As Boolean
    OpIsRunning = mOpRunning
End Property

Friend Property Get OpLastError() As String
    OpLastError = mOpLastError
End Property

Friend Sub RequestOpCancel()
    RequireRole ROLE_ASYNC_OP, "RequestOpCancel"
    If mOpCts Is Nothing Then
        RaiseError ERROR_INVALID_OPERATION, "CancelAsync", _
            "Async op '" & mComponentId & "' was not built with WithCancellation."
    End If
    mOpCts.Cancel
End Sub

Friend Sub TickOp(ByVal ownerApp As ReDimUI)
    If Not mOpRunning Then Exit Sub
    mOpTask.AdvanceTask
    If Not mOpTask.IsCompleted Then Exit Sub
    FinishOp ownerApp
End Sub

Private Sub FinishOp(ByVal ownerApp As ReDimUI)
    Dim outcomeProc As String
    Dim outcomeState As String

    mOpRunning = False
    If mOpTask.IsFaulted Then
        outcomeState = "failed"
        outcomeProc = mOpFailProc
        On Error Resume Next
        mOpLastError = mOpTask.Exception.Message
        On Error GoTo 0
    ElseIf mOpTask.IsCanceled Then
        outcomeState = "canceled"
        outcomeProc = mOpCancelProc
    Else
        outcomeState = "done"
        outcomeProc = mOpDoneProc
    End If
    SetOpControlsBusy ownerApp, False
    If LenB(mOpStatusKey) > 0 Then ownerApp.SetState mOpStatusKey, outcomeState
    If mOpTask.IsFaulted Then
        ownerApp.ReportAppError "ReDim async op '" & mComponentId & _
            "' failed: " & mOpLastError
    End If
    If LenB(outcomeProc) > 0 Then
        On Error Resume Next
        Application.Run outcomeProc
        On Error GoTo 0
    End If
End Sub

Private Sub SetOpControlsBusy(ByVal ownerApp As ReDimUI, ByVal busyValue As Boolean)
    Dim componentId As Variant
    Dim componentValue As ReDimUI
    Dim spinnerValue As ReDimUI

    For Each componentId In mOpDisableIds
        If ownerApp.TryGetComponent(CStr(componentId), componentValue) Then
            componentValue.SetBusy busyValue
        End If
    Next componentId
    If LenB(mOpSpinnerId) > 0 Then
        If ownerApp.TryGetComponent(mOpSpinnerId, spinnerValue) Then
            spinnerValue.Visible busyValue
        End If
    End If
End Sub

Private Sub RequireNotRunning(ByVal operation As String)
    If mOpRunning Then
        RaiseError ERROR_INVALID_OPERATION, operation, _
            "Async op '" & mComponentId & "' is already running."
    End If
    If mJobRunning Then
        RaiseError ERROR_INVALID_OPERATION, operation, _
            "Job '" & mComponentId & "' is already running."
    End If
End Sub

' =====================================================================
' Jobs: chunked CPU-bound work driven inside a per-tick budget
' =====================================================================

Friend Sub ConfigureJob(ByVal ownerAppId As String, ByVal jobId As String)
    mRole = ROLE_JOB
    mOwnerAppId = ownerAppId
    mComponentId = jobId
    mJobBudgetMs = 15
End Sub

' The step procedure is a zero-argument Function returning True when the job
' has finished. It runs repeatedly inside each tick until the budget expires.
Public Function Steps(ByVal procName As String) As ReDimUI
    RequireRole ROLE_JOB, "Steps"
    RequireNotRunning "Steps"
    mJobStepProc = procName
    Set Steps = Me
End Function

Public Function BudgetMs(ByVal budget As Long) As ReDimUI
    RequireRole ROLE_JOB, "BudgetMs"
    If budget < 1 Then budget = 1
    mJobBudgetMs = budget
    mJobBudgetExplicit = True
    mJobPaceMs = 0
    Set BudgetMs = Me
End Function

' Paced mode: at most one step per interval, for game loops and animations
' where cadence matters more than throughput. Callable while running so a
' game can change speed live.
Public Function PacedMs(ByVal interval As Long) As ReDimUI
    RequireRole ROLE_JOB, "PacedMs"
    If interval < 1 Then interval = 1
    mJobPaceMs = interval
    Set PacedMs = Me
End Function

Public Function JobOnDone(ByVal procName As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnDone"
    mJobDoneProc = procName
    Set JobOnDone = Me
End Function

Public Function JobOnFail(ByVal procName As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnFail"
    mJobFailProc = procName
    Set JobOnFail = Me
End Function

Public Function JobOnCancel(ByVal procName As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnCancel"
    mJobCancelProc = procName
    Set JobOnCancel = Me
End Function

Public Sub StartJob()
    RequireRole ROLE_JOB, "StartJob"
    RequireNotRunning "StartJob"
    If LenB(mJobStepProc) = 0 Then
        RaiseError ERROR_INVALID_OPERATION, "StartJob", _
            "Job '" & mComponentId & "' needs Steps first."
    End If
    mJobCancelRequested = False
    mJobRunning = True
    ReDimUI.RequestPump
End Sub

Friend Property Get JobIsRunning() As Boolean
    JobIsRunning = mJobRunning
End Property

Friend Sub RequestJobCancel()
    RequireRole ROLE_JOB, "RequestJobCancel"
    mJobCancelRequested = True
End Sub

Friend Sub TickJob( _
    ByVal ownerApp As ReDimUI, _
    ByVal nowMs As LongLong, _
    ByVal runWork As Boolean, _
    ByVal animating As Boolean _
)
    Dim stepDone As Variant
    Dim effectiveBudget As Long

    If Not mJobRunning Then Exit Sub
    If mJobCancelRequested Then
        FinishJob ownerApp, mJobCancelProc
        Exit Sub
    End If
    On Error GoTo StepFailed
    If mJobPaceMs > 0 Then
        If nowMs - mJobLastStepMs < CLngLng(mJobPaceMs) Then Exit Sub
        mJobLastStepMs = nowMs
        stepDone = Application.Run(mJobStepProc)
        If CBool(stepDone) Then FinishJob ownerApp, mJobDoneProc
        On Error GoTo 0
        Exit Sub
    End If
    If Not runWork Then
        On Error GoTo 0
        Exit Sub
    End If
    ' While something animates, a default budget yields to the frame rate;
    ' an explicit BudgetMs is always honored as given. Budget timing uses
    ' the precise clock so sub-16 ms budgets mean what they say.
    effectiveBudget = mJobBudgetMs
    If animating And Not mJobBudgetExplicit Then effectiveBudget = 8
    Dim startedPreciseMs As Double
    startedPreciseMs = PreciseNowMs()
    Do
        stepDone = Application.Run(mJobStepProc)
        If CBool(stepDone) Then
            FinishJob ownerApp, mJobDoneProc
            Exit Sub
        End If
        If mJobCancelRequested Then
            FinishJob ownerApp, mJobCancelProc
            Exit Sub
        End If
    Loop While PreciseNowMs() - startedPreciseMs < effectiveBudget
    On Error GoTo 0
    Exit Sub

StepFailed:
    Dim failure As String
    failure = "ReDim job '" & mComponentId & "' failed: " & Err.Description
    On Error GoTo 0
    FinishJob ownerApp, mJobFailProc
    ownerApp.ReportAppError failure
End Sub

Private Sub FinishJob(ByVal ownerApp As ReDimUI, ByVal outcomeProc As String)
    mJobRunning = False
    mJobCancelRequested = False
    If LenB(outcomeProc) > 0 Then
        On Error Resume Next
        Application.Run outcomeProc
        On Error GoTo 0
    End If
End Sub

' =====================================================================
' Theme configuration
' =====================================================================

Friend Sub ConfigureTheme( _
    ByVal primaryColor As Long, ByVal onPrimaryColor As Long, _
    ByVal surfaceColor As Long, ByVal onSurfaceColor As Long, _
    ByVal mutedColor As Long, ByVal onMutedColor As Long, _
    ByVal successColor As Long, ByVal dangerColor As Long, _
    ByVal borderColor As Long, ByVal canvasColor As Long, _
    ByVal fontName As String, ByVal baseFontSize As Double _
)
    mRole = ROLE_THEME
    mThemePrimary = primaryColor
    mThemeOnPrimary = onPrimaryColor
    mThemeSurface = surfaceColor
    mThemeOnSurface = onSurfaceColor
    mThemeMuted = mutedColor
    mThemeOnMuted = onMutedColor
    mThemeSuccess = successColor
    mThemeDanger = dangerColor
    mThemeBorder = borderColor
    mThemeCanvas = canvasColor
    mThemeFontName = fontName
    mThemeFontSize = baseFontSize
End Sub

Public Property Get PrimaryColor() As Long
    PrimaryColor = mThemePrimary
End Property

Public Property Get OnPrimaryColor() As Long
    OnPrimaryColor = mThemeOnPrimary
End Property

Public Property Get SurfaceColor() As Long
    SurfaceColor = mThemeSurface
End Property

Public Property Get OnSurfaceColor() As Long
    OnSurfaceColor = mThemeOnSurface
End Property

Public Property Get MutedColor() As Long
    MutedColor = mThemeMuted
End Property

Public Property Get OnMutedColor() As Long
    OnMutedColor = mThemeOnMuted
End Property

Public Property Get SuccessColor() As Long
    SuccessColor = mThemeSuccess
End Property

Public Property Get DangerColor() As Long
    DangerColor = mThemeDanger
End Property

Public Property Get BorderColor() As Long
    BorderColor = mThemeBorder
End Property

Public Property Get CanvasColor() As Long
    CanvasColor = mThemeCanvas
End Property

Public Property Get FontName() As String
    FontName = mThemeFontName
End Property

Public Property Get BaseFontSize() As Double
    BaseFontSize = mThemeFontSize
End Property

Public Function WithPrimary(ByVal color As Long, ByVal onColor As Long) As ReDimUI
    RequireRole ROLE_THEME, "WithPrimary"
    mThemePrimary = color
    mThemeOnPrimary = onColor
    Set WithPrimary = Me
End Function

Public Function WithFont(ByVal name As String, ByVal size As Double) As ReDimUI
    RequireRole ROLE_THEME, "WithFont"
    mThemeFontName = name
    mThemeFontSize = size
    Set WithFont = Me
End Function

' =====================================================================
' Shared helpers
' =====================================================================

Friend Property Get InstancePointer() As LongPtr
    InstancePointer = ObjPtr(Me)
End Property

Private Function IsFactory(ByVal instance As ReDimUI) As Boolean
    IsFactory = (ObjPtr(instance) = ReDimUI.InstancePointer)
End Function

Private Sub EnsureFactory()
    If mApps Is Nothing Then Set mApps = New Collection
End Sub

Private Function ParseShapeName( _
    ByVal shapeName As String, _
    ByRef targetAppId As String, _
    ByRef targetComponentId As String, _
    ByRef targetPart As String _
) As Boolean
    Dim remainder As String
    Dim separatorAt As Long
    Dim partAt As Long

    targetPart = vbNullString
    If Left$(shapeName, Len(SHAPE_PREFIX)) <> SHAPE_PREFIX Then Exit Function
    remainder = Mid$(shapeName, Len(SHAPE_PREFIX) + 1)
    separatorAt = InStr(1, remainder, "_")
    If separatorAt = 0 Then Exit Function
    targetAppId = Left$(remainder, separatorAt - 1)
    targetComponentId = Mid$(remainder, separatorAt + 1)
    partAt = InStr(1, targetComponentId, PART_SEPARATOR)
    If partAt > 0 Then
        targetPart = Mid$(targetComponentId, partAt + Len(PART_SEPARATOR))
        targetComponentId = Left$(targetComponentId, partAt - 1)
    End If
    ParseShapeName = (LenB(targetAppId) > 0 And LenB(targetComponentId) > 0)
End Function

Private Sub RequireIdentifier( _
    ByVal candidate As String, _
    ByVal allowUnderscore As Boolean, _
    ByVal operation As String, _
    ByVal argumentName As String _
)
    Dim position As Long
    Dim code As Integer

    If LenB(candidate) = 0 Then
        RaiseError ERROR_INVALID_ARGUMENT, operation, argumentName & " is required."
    End If
    For position = 1 To Len(candidate)
        code = Asc(UCase$(Mid$(candidate, position, 1)))
        If Not ((code >= 65 And code <= 90) Or (code >= 48 And code <= 57) _
            Or (allowUnderscore And code = 95)) Then
            RaiseError ERROR_INVALID_ARGUMENT, operation, _
                argumentName & " may contain only letters, digits" & _
                IIf(allowUnderscore, ", and single underscores.", ".")
        End If
    Next position
    If allowUnderscore Then
        If InStr(1, candidate, PART_SEPARATOR) > 0 Then
            RaiseError ERROR_INVALID_ARGUMENT, operation, _
                argumentName & " may not contain a double underscore."
        End If
    End If
End Sub

Private Sub RequireRole(ByVal expectedRole As Long, ByVal operation As String)
    If mRole <> expectedRole Then
        RaiseError ERROR_INVALID_OPERATION, operation, _
            operation & " is not valid on this ReDimUI value."
    End If
End Sub

Private Sub RaiseError( _
    ByVal errorNumber As Long, _
    ByVal operation As String, _
    ByVal message As String _
)
    Err.Raise errorNumber, "ReDimUI." & operation, message
End Sub

Private Function BoolToMso(ByVal value As Boolean) As Long
    If value Then BoolToMso = msoTrue Else BoolToMso = msoFalse
End Function

' Millisecond clock from QueryPerformanceCounter: sub-millisecond precision
' regardless of the system timer resolution. Deltas only; its epoch is not
' comparable to GetTickCount64.
Private Function PreciseNowMs() As Double
    Static frequency As LongLong
    Dim counter As LongLong

    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    PreciseNowMs = counter * 1000# / frequency
End Function

Private Sub AssignVariant(ByRef target As Variant, ByVal source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub

Private Function TryGetItem( _
    ByVal source As Collection, _
    ByVal key As String, _
    ByRef result As Variant _
) As Boolean
    On Error GoTo Missing
    AssignVariant result, source.Item(key)
    TryGetItem = True
Missing:
End Function

Private Function TryGetObject( _
    ByVal source As Collection, _
    ByVal key As String, _
    ByRef result As ReDimUI _
) As Boolean
    On Error GoTo Missing
    Set result = source.Item(key)
    TryGetObject = True
Missing:
End Function

Private Sub ReplaceItem( _
    ByVal source As Collection, _
    ByVal key As String, _
    ByVal newValue As Variant _
)
    On Error Resume Next
    source.Remove key
    On Error GoTo 0
    source.Add newValue, key
End Sub
