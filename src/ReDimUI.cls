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
' ReDim 0.20.0 (2026-09-22)
' https://github.com/WilliamSmithEdward/ReDim
'
' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

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
    ByRef ticksPerSecond As LongLong) As Long
Private Type RdxScreenPoint
    pixelX As Long
    pixelY As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
    ByRef screenPoint As RdxScreenPoint) As Long
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
Private Const KIND_TRANSFER As Long = 19
Private Const KIND_CHECKLIST As Long = 20
Private Const KIND_IMAGE As Long = 21

' Option rows a SelectBox or ComboBox drop list draws at once unless
' ListRows says otherwise; longer lists window behind pager rows.
Private Const DEFAULT_LIST_ROWS As Long = 8

' Sentinel for "the source could not be loaded"; contains characters no
' real path can, so it never collides with one.
Private Const IMAGE_MISSING As String = "?missing?"

Private Const VARIANT_PRIMARY As Long = 0
Private Const VARIANT_SECONDARY As Long = 1
Private Const VARIANT_SUCCESS As Long = 2
Private Const VARIANT_DANGER As Long = 3
Private Const VARIANT_NEUTRAL As Long = 4

Private Const REL_NONE As Long = 0
Private Const REL_BELOW As Long = 1
Private Const REL_RIGHT_OF As Long = 2
Private Const REDIM_VERSION As String = "0.20.0"

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
' Test seam: the next real tick treats the field-watch press edge as
' fired, so in-tick blur paths are testable without physical input.
Private mForcePressEdge As Boolean
' Guarded-swallow observability: every silently trapped tick or key
' dispatch fault counts here, so tests can prove a run was clean.
Private mTickFaults As Long
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
Private mToastRailLeft As Double
Private mToastRailTop As Double
Private mToastRailStored As Boolean
Private mHotKeys As Collection
Private mSurfaceProtected As Boolean
Private mIsWindow As Boolean
Private mOnShowProc As String
Private mOnHideProc As String
Private mCanvasPrepared As Boolean
Private mCanvasPainted As Boolean
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
Private mToastExiting As Boolean
Private mExitRemaining As Double
Private mItems As Collection
' Transfer list: the right-hand (chosen) side and its selections. mItems
' carries the left-hand (available) side, so the item APIs work unchanged.
' Selections are sets of row indices: Excel reserves Ctrl+click on macro
' shapes for selecting the shape itself (the macro never runs), so rows
' multi-select by plain click-to-toggle instead of a modifier.
Private mChosenItems As Collection
Private mSelLeft As Collection
Private mSelRight As Collection
' Checkbox list: the set of checked row indices, and whether the
' select-all header row renders (on by default).
Private mCheckedSet As Collection
Private mSelectAllRow As Boolean
' Image control: the source file path, its optional state binding, and
' the last source actually painted (so files load once, not per flush).
Private mImageSource As String
Private mBindSourceKey As String
Private mAppliedImageSource As String
' Multi-line float TextInput: Enter inserts a newline instead of
' committing; Tab, Ctrl+Enter, and clicking away commit.
Private mMultiLine As Boolean
' Diff caches for the typing hot path: the float focus ring state and
' the drop list signature last painted, plus the composite repaint
' signatures for the transfer list and checkbox list.
Private mAppliedFocusRing As Boolean
Private mAppliedListSig As String
Private mAppliedTransferSig As String
Private mAppliedCheckSig As String
Private mAppliedProgressSig As String
Private mAppliedCellFill As Long
' The caret index: characters before the caret, 0 through Len(mText).
Private mCaretPos As Long
' Windowed long lists: a SelectBox or ComboBox drop list shows at most
' mListRows of mListMatches (item indexes in list order) from
' mListOffset, with mListHighlight as the combo's keyboard cursor
' (1-based into the matches, 0 for none). A pristine combo list is
' unfiltered while its text names an item, until the user edits it.
' Transfer panels scroll through per-side offsets.
Private mListRows As Long
Private mListMatches As Collection
Private mFilterPristine As Boolean
Private mListOffset As Long
Private mListHighlight As Long
Private mLeftOffset As Long
Private mRightOffset As Long
Private mCaptionLeft As String
Private mCaptionRight As String
Private mAppliedLeftRows As Long
Private mAppliedRightRows As Long
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
Private mFieldTextAtFocus As String
Private mFieldSelAtFocus As String
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

Public Function Mount( _
    ByVal hostSheet As Worksheet, _
    ByVal targetAppId As String _
) As ReDimUI
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        Set Mount = ReDimUI.Mount(hostSheet, targetAppId)
        Exit Function
    End If
    RequireIdentifier targetAppId, False, "Mount", "appId"
    If hostSheet Is Nothing Then
        RaiseError ERROR_INVALID_ARGUMENT, "Mount", "Mount requires a worksheet."
    End If
    EnsureFactory
    If mExcelApp Is Nothing Then Set mExcelApp = Application
    Set appValue = FindApp(targetAppId)
    If appValue Is Nothing Then
        Set appValue = New ReDimUI
        appValue.ConfigureApp targetAppId, hostSheet
        mApps.Add appValue, targetAppId
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

Private Function FindApp(ByVal targetAppId As String) As ReDimUI
    On Error Resume Next
    Set FindApp = mApps.Item(targetAppId)
    On Error GoTo 0
End Function

Public Property Get App(ByVal targetAppId As String) As ReDimUI
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        Set App = ReDimUI.App(targetAppId)
        Exit Property
    End If
    EnsureFactory
    If Not TryGetObject(mApps, targetAppId, appValue) Then
        RaiseError ERROR_NOT_FOUND, "App", _
            "No mounted app is named '" & targetAppId & "'. Call Mount first."
    End If
    Set App = appValue
End Property

Public Property Get HasApp(ByVal targetAppId As String) As Boolean
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        HasApp = ReDimUI.HasApp(targetAppId)
        Exit Property
    End If
    EnsureFactory
    HasApp = TryGetObject(mApps, targetAppId, appValue)
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
Friend Sub DispatchShape(ByVal clickedName As String)
    Dim appValue As ReDimUI
    Dim componentValue As ReDimUI
    Dim targetAppId As String
    Dim targetComponentId As String
    Dim targetPart As String

    If Not IsFactory(Me) Then
        ReDimUI.DispatchShape clickedName
        Exit Sub
    End If
    EnsureFactory
    If Not ParseShapeName(clickedName, targetAppId, targetComponentId, targetPart) Then
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

' Test seam: consumed by the first TickApp that polls after it is set.
' Single-app scenarios only; with several apps mounted the first app in
' registry order consumes the edge.
Friend Sub ForcePressEdge()
    If Not IsFactory(Me) Then
        ReDimUI.ForcePressEdge
        Exit Sub
    End If
    mForcePressEdge = True
End Sub

Friend Function ConsumeForcedPressEdge() As Boolean
    If Not IsFactory(Me) Then
        ConsumeForcedPressEdge = ReDimUI.ConsumeForcedPressEdge
        Exit Function
    End If
    ConsumeForcedPressEdge = mForcePressEdge
    mForcePressEdge = False
End Function

' The count of silently trapped faults since the last reset. The pump
' and key dispatch never surface errors to the user; this counter is
' how tests (and diagnostics) prove that silence was earned.
Friend Sub NoteTickFault()
    If Not IsFactory(Me) Then
        ReDimUI.NoteTickFault
        Exit Sub
    End If
    mTickFaults = mTickFaults + 1
End Sub

Public Property Get TickFaultCount() As Long
    If Not IsFactory(Me) Then
        TickFaultCount = ReDimUI.TickFaultCount
        Exit Property
    End If
    TickFaultCount = mTickFaults
End Property

Public Sub ResetTickFaults()
    If Not IsFactory(Me) Then
        ReDimUI.ResetTickFaults
        Exit Sub
    End If
    mTickFaults = 0
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
    ByVal targetAppId As String, _
    ByVal targetComponentId As String _
) As Boolean
    If Not IsFactory(Me) Then
        IsComponentFocused = ReDimUI.IsComponentFocused(targetAppId, targetComponentId)
        Exit Function
    End If
    IsComponentFocused = (mFocusAppId = targetAppId _
        And mFocusComponentId = targetComponentId)
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
    For Each appValue In mApps
        On Error Resume Next
        Err.Clear
        appValue.TickApp nowMs, elapsedMs, runWork
        If Err.Number <> 0 Then
            Err.Clear
            NoteTickFault
        End If
        On Error GoTo 0
    Next appValue
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
Public Sub PinPumpCursor(ByVal pinOn As Boolean)
    RdxSetCursorPin pinOn
End Sub

' Deterministic tests turn the wall-clock timer off and drive PumpOnce
' themselves. Production leaves it on.
Public Sub AutoPump(ByVal pumpOn As Boolean)
    If Not IsFactory(Me) Then
        ReDimUI.AutoPump pumpOn
        Exit Sub
    End If
    mAutoPumpOff = Not pumpOn
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
Public Sub Navigate(ByVal targetAppId As String)
    If Not IsFactory(Me) Then
        ReDimUI.Navigate targetAppId
        Exit Sub
    End If
    NavigateCore targetAppId, True
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

Private Sub NavigateCore(ByVal targetAppId As String, ByVal pushCurrent As Boolean)
    Dim targetApp As ReDimUI
    Dim priorApp As ReDimUI
    Dim windowId As Variant
    Dim windowApp As ReDimUI
    Dim previousId As String

    EnsureFactory
    Set targetApp = App(targetAppId)
    If Not targetApp.IsWindow Then
        RaiseError ERROR_INVALID_OPERATION, "Navigate", _
            "App '" & targetAppId & "' is not a window. Call AsWindow on it first."
    End If
    If mActiveWindowId = targetAppId Then Exit Sub
    previousId = mActiveWindowId
    targetApp.Sheet.Visible = xlSheetVisible
    targetApp.Sheet.Activate
    targetApp.ReassertCanvas
    For Each windowId In mWindows
        If CStr(windowId) <> targetAppId Then
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
        If TryGetObject(mApps, previousId, priorApp) Then priorApp.InvokeHide
    End If
    mActiveWindowId = targetAppId
    For Each windowId In mWindows
        If TryGetObject(mApps, CStr(windowId), windowApp) Then
            windowApp.RefreshNavBar
        End If
    Next windowId
    ' The newly shown window may carry pump demand (sliders, spinners).
    RequestPump
    targetApp.InvokeShow
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

Friend Sub RegisterWindow(ByVal targetAppId As String)
    Dim windowId As Variant

    If Not IsFactory(Me) Then
        ReDimUI.RegisterWindow targetAppId
        Exit Sub
    End If
    EnsureFactory
    If mWindows Is Nothing Then Set mWindows = New Collection
    For Each windowId In mWindows
        If CStr(windowId) = targetAppId Then Exit Sub
    Next windowId
    mWindows.Add targetAppId
End Sub

Friend Sub ForgetWindow(ByVal targetAppId As String)
    Dim idx As Long

    If Not IsFactory(Me) Then
        ReDimUI.ForgetWindow targetAppId
        Exit Sub
    End If
    EnsureFactory
    If mWindows Is Nothing Then Exit Sub
    For idx = mWindows.Count To 1 Step -1
        If CStr(mWindows.Item(idx)) = targetAppId Then mWindows.Remove idx
    Next idx
    If Not mNavStack Is Nothing Then
        For idx = mNavStack.Count To 1 Step -1
            If CStr(mNavStack.Item(idx)) = targetAppId Then
                mNavStack.Remove idx
            End If
        Next idx
    End If
    If mActiveWindowId = targetAppId Then mActiveWindowId = vbNullString
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

Friend Sub ConfigureApp(ByVal targetAppId As String, ByVal hostSheet As Worksheet)
    mRole = ROLE_APP
    mAppId = targetAppId
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

Public Function Button(ByVal targetComponentId As String) As ReDimUI
    Set Button = GetOrCreateComponent(targetComponentId, KIND_BUTTON)
End Function

Public Function Label(ByVal targetComponentId As String) As ReDimUI
    Set Label = GetOrCreateComponent(targetComponentId, KIND_LABEL)
End Function

Public Function Card(ByVal targetComponentId As String) As ReDimUI
    Set Card = GetOrCreateComponent(targetComponentId, KIND_CARD)
End Function

Public Function Spinner(ByVal targetComponentId As String) As ReDimUI
    Set Spinner = GetOrCreateComponent(targetComponentId, KIND_SPINNER)
End Function

Public Function ProgressBar(ByVal targetComponentId As String) As ReDimUI
    Set ProgressBar = GetOrCreateComponent(targetComponentId, KIND_PROGRESS)
End Function

Public Function Toggle(ByVal targetComponentId As String) As ReDimUI
    Set Toggle = GetOrCreateComponent(targetComponentId, KIND_TOGGLE)
End Function

' Drawn picker: face, caret, and option list in the app font. Text is
' the placeholder shown while nothing is selected, and Value(0) returns
' to it. The list windows to ListRows rows (eight by default).
Public Function SelectBox(ByVal targetComponentId As String) As ReDimUI
    Set SelectBox = GetOrCreateComponent(targetComponentId, KIND_SELECT)
End Function

' Drawn checkbox: themed box, check glyph, and caption. Click the box or
' the caption to toggle.
Public Function TickBox(ByVal targetComponentId As String) As ReDimUI
    Set TickBox = GetOrCreateComponent(targetComponentId, KIND_TICKBOX)
End Function

' Drawn dual listbox (transfer list): two panels with move buttons between
' them. Items feeds the available side, ChosenFrom the chosen side; click a
' row to select it, then move one (>, <) or all (>>, <<). WritesTo carries
' the chosen items joined with ", "; OnChange fires once per user transfer.
Public Function TransferList(ByVal targetComponentId As String) As ReDimUI
    Set TransferList = GetOrCreateComponent(targetComponentId, KIND_TRANSFER)
End Function

' Drawn checkbox list: one box-and-caption row per item, any number
' checked, with a select-all header row (tri-state box: empty, check,
' dash for mixed) on by default. Box and caption both toggle their row.
' WritesTo carries the checked items joined with ", "; OnChange fires
' once per user toggle, including the select-all bulk toggle.
Public Function CheckList(ByVal targetComponentId As String) As ReDimUI
    Set CheckList = GetOrCreateComponent(targetComponentId, KIND_CHECKLIST)
End Function

' Drawn image: a rounded rectangle whose fill is the picture, so corners,
' geometry, adoption, and clicks behave like every other control and the
' picture embeds in the workbook. Source takes a file path (no URLs);
' the image stretches to the declared rectangle; a missing file renders
' a themed placeholder, but an already-embedded picture is kept even if
' its source file has since gone away.
Public Function Image(ByVal targetComponentId As String) As ReDimUI
    Set Image = GetOrCreateComponent(targetComponentId, KIND_IMAGE)
End Function

' Drawn single-select option rows. Items supplies the rows, Value the
' selected index, and WritesTo receives the selected item's text.
Public Function RadioGroup(ByVal targetComponentId As String) As ReDimUI
    Set RadioGroup = GetOrCreateComponent(targetComponentId, KIND_RADIO)
End Function

' Drawn numeric control: minus and plus around a value face, honoring
' SliderRange. Shapes cannot report click coordinates, so a drag-thumb
' slider cannot be drawn honestly; stepped entry can.
Public Function Stepper(ByVal targetComponentId As String) As ReDimUI
    Set Stepper = GetOrCreateComponent(targetComponentId, KIND_STEPPER)
End Function

' Drawn slider, click-to-set: the click position is mapped back into sheet
' points by inverting PointsToScreenPixels, then snapped to SliderRange.
' Loop-free by design; nothing in ReDim may block the pump.
Public Function SlideBar(ByVal targetComponentId As String) As ReDimUI
    Set SlideBar = GetOrCreateComponent(targetComponentId, KIND_SLIDEBAR)
End Function

' Editable combo: a text face with a caret and a drawn, windowed option
' list. A float combo (AtRect) filters live as keys arrive. A cell-backed
' combo (At) filters at the moments Excel grants, because edit mode
' pauses all VBA: committing with Enter auto-suggests matches (an exact
' match simply takes it). A picked item reopens onto the whole list
' until the text is edited. Free text is always allowed.
Public Function ComboBox(ByVal targetComponentId As String) As ReDimUI
    Set ComboBox = GetOrCreateComponent(targetComponentId, KIND_COMBO)
End Function

Public Function TextInput(ByVal targetComponentId As String) As ReDimUI
    Set TextInput = GetOrCreateComponent(targetComponentId, KIND_INPUT)
End Function

' Transient notification card. Expires on the pump after ttlMs, or when
' clicked.
Public Function Toast( _
    ByVal messageText As String, _
    Optional ByVal ttlMs As Long = 3000 _
) As ReDimUI
    Dim toastValue As ReDimUI

    RequireRole ROLE_APP, "Toast"
    mToastCounter = mToastCounter + 1
    Set toastValue = GetOrCreateComponent("toast_" & mToastCounter, KIND_TOAST)
    toastValue.ConfigureToast messageText, ttlMs, NextToastSlot()
    NotifyDirty toastValue.ComponentId
    If mRendered Then FlushDirty
    ReDimUI.RequestPump
    Set Toast = toastValue
End Function

' Pins the toast tray's top-left to a range anchor instead of the default
' rail beside the app content.
Public Function ToastTray(ByVal rangeAddress As String) As ReDimUI
    RequireRole ROLE_APP, "ToastTray"
    mToastTrayAnchor = rangeAddress
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

' The tray origin is sticky while any toast is alive: survivors and
' newcomers must share one column, so the origin computed by the toast
' that opened the tray is reused until the tray empties. Recomputing per
' spawn let scroll, zoom, and content-bounds shifts move the rail
' mid-stream, landing a new toast on top of live ones. The claimant is
' excluded from the liveness scan because it is already registered by
' the time it asks.
Friend Function ActiveToastRail( _
    ByVal claimantId As String, _
    ByRef originLeft As Double, _
    ByRef originTop As Double _
) As Boolean
    Dim componentValue As ReDimUI

    If Not mToastRailStored Then Exit Function
    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then
            If componentValue.ComponentId <> claimantId Then
                originLeft = mToastRailLeft
                originTop = mToastRailTop
                ActiveToastRail = True
                Exit Function
            End If
        End If
    Next componentValue
End Function

Friend Sub RecordToastRail( _
    ByVal originLeft As Double, _
    ByVal originTop As Double _
)
    mToastRailLeft = originLeft
    mToastRailTop = originTop
    mToastRailStored = True
End Sub

' The lowest visual position any live toast currently occupies, glides
' included. -1 when the tray is otherwise empty.
Friend Function MaxToastVisualTop(ByVal claimantId As String) As Double
    Dim componentValue As ReDimUI

    MaxToastVisualTop = -1
    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then
            If componentValue.ComponentId <> claimantId Then
                If componentValue.ToastVisualTop > MaxToastVisualTop Then
                    MaxToastVisualTop = componentValue.ToastVisualTop
                End If
            End If
        End If
    Next componentValue
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
    ByVal titleText As String, _
    ByVal messageText As String, _
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
        .Text(titleText & vbLf & vbLf & messageText).Visible True
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

Private Sub HideIfExists(ByVal targetComponentId As String)
    Dim componentValue As ReDimUI

    If TryGetComponent(targetComponentId, componentValue) Then
        componentValue.Visible False
    End If
End Sub

Private Sub RaiseModalZOrder()
    Dim modalId As Variant
    Dim componentValue As ReDimUI

    For Each modalId In Array("mdl_ov", "mdl_card", "mdl_ok", "mdl_cancel")
        If TryGetComponent(CStr(modalId), componentValue) Then
            On Error Resume Next
            mSheet.Shapes(componentValue.ShapeName).ZOrder msoBringToFront
            On Error GoTo 0
        End If
    Next modalId
End Sub

Friend Sub ResolveViewport( _
    ByRef viewLeft As Double, _
    ByRef viewTop As Double, _
    ByRef viewWidth As Double, _
    ByRef viewHeight As Double _
)
    Dim shownRange As Range

    On Error Resume Next
    If ActiveWindow Is Nothing Then
    ElseIf mSheet Is ActiveWindow.ActiveSheet Then
        Set shownRange = ActiveWindow.VisibleRange
    End If
    On Error GoTo 0
    If shownRange Is Nothing Then Set shownRange = mSheet.Range("A1:P36")
    viewLeft = shownRange.Left
    viewTop = shownRange.Top
    viewWidth = shownRange.Width
    viewHeight = shownRange.Height
End Sub

' Paints the sheet background and hides gridlines so shapes sit on a clean
' canvas. Gridlines are window state, so they change only while this sheet
' is active.
Public Function PrepareCanvas( _
    Optional ByVal hideGridlines As Boolean = True, _
    Optional ByVal paintBackground As Boolean = True _
) As ReDimUI
    RequireRole ROLE_APP, "PrepareCanvas"
    If paintBackground Then
        mSheet.Cells.Interior.Color = mTheme.CanvasColor
        mCanvasPainted = True
    End If
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
    ByRef extentLeft As Double, _
    ByRef extentTop As Double, _
    ByRef extentRight As Double _
) As Boolean
    Dim componentValue As ReDimUI

    For Each componentValue In mComponents
        If componentValue.CountsForBounds Then
            If Not ContentBounds Then
                extentLeft = componentValue.BoundsLeft
                extentTop = componentValue.BoundsTop
                extentRight = componentValue.BoundsRight
                ContentBounds = True
            Else
                If componentValue.BoundsLeft < extentLeft Then
                    extentLeft = componentValue.BoundsLeft
                End If
                If componentValue.BoundsTop < extentTop Then
                    extentTop = componentValue.BoundsTop
                End If
                If componentValue.BoundsRight > extentRight Then
                    extentRight = componentValue.BoundsRight
                End If
            End If
        End If
    Next componentValue
End Function

Friend Function GetOrCreateComponent( _
    ByVal targetComponentId As String, _
    ByVal kindCode As Long _
) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Component"
    RequireIdentifier targetComponentId, True, "Component", "componentId"
    If TryGetComponent(targetComponentId, componentValue) Then
        If componentValue.ComponentKind <> kindCode Then
            RaiseError ERROR_INVALID_OPERATION, "Component", _
                "Component '" & targetComponentId & _
                "' already exists with a different kind."
        End If
        Set GetOrCreateComponent = componentValue
        Exit Function
    End If
    Set componentValue = New ReDimUI
    componentValue.ConfigureComponent mAppId, targetComponentId, kindCode
    mComponents.Add componentValue, targetComponentId
    mComponentOrder.Add targetComponentId
    Set GetOrCreateComponent = componentValue
End Function

Friend Function TryGetComponent( _
    ByVal targetComponentId As String, _
    ByRef found As ReDimUI _
) As Boolean
    TryGetComponent = TryGetObject(mComponents, targetComponentId, found)
End Function

Friend Function PeekComponent(ByVal targetComponentId As String) As ReDimUI
    Set PeekComponent = FindKeyed(mComponents, targetComponentId)
End Function

Public Property Get Component(ByVal targetComponentId As String) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Component"
    If Not TryGetComponent(targetComponentId, componentValue) Then
        RaiseError ERROR_NOT_FOUND, "Component", _
            "No component is named '" & targetComponentId & "'."
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
    MarkBoundComponentsDirty stateKey
    InvokeStateHandlers stateKey
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

' Sets the key only when it has no value yet, so setup code stays
' idempotent: re-running a build never clobbers values already in play.
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

' The state store is deliberately in-memory only. Durability is the host
' application's concern: read the store with State/StateOrDefault and save
' wherever fits (a hidden sheet, workbook names, a file), then reseed on
' build with SetStateDefault. ROneCOne.Json serializes and parses if JSON
' is the format of choice.

' Registers a zero-argument workbook procedure invoked after the key changes.
Public Sub OnStateChanged(ByVal stateKey As String, ByVal handlerProc As String)
    Dim handlers As Collection

    RequireRole ROLE_APP, "OnStateChanged"
    Set handlers = FindHandlers(stateKey)
    If handlers Is Nothing Then
        Set handlers = New Collection
        mStateHandlers.Add handlers, stateKey
    End If
    handlers.Add handlerProc
End Sub

Private Sub InvokeStateHandlers(ByVal stateKey As String)
    Dim handlers As Collection
    Dim handlerProc As Variant

    Set handlers = FindHandlers(stateKey)
    If handlers Is Nothing Then Exit Sub
    For Each handlerProc In handlers
        Application.Run CStr(handlerProc)
    Next handlerProc
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
    Dim orderedId As Variant

    RequireRole ROLE_APP, "Render"
    For Each orderedId In mComponentOrder
        NotifyDirty CStr(orderedId)
    Next orderedId
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
    Dim candidateShape As Shape
    Dim orphanNames As Collection
    Dim orphanName As Variant
    Dim namePrefix As String
    Dim remainder As String
    Dim orphanComponentId As String
    Dim partAt As Long
    Dim componentValue As ReDimUI

    Set orphanNames = New Collection
    namePrefix = SHAPE_PREFIX & mAppId & "_"
    For Each candidateShape In mSheet.Shapes
        If Left$(candidateShape.Name, Len(namePrefix)) = namePrefix Then
            remainder = Mid$(candidateShape.Name, Len(namePrefix) + 1)
            partAt = InStr(1, remainder, PART_SEPARATOR)
            If partAt > 0 Then
                orphanComponentId = Left$(remainder, partAt - 1)
            Else
                orphanComponentId = remainder
            End If
            If Not TryGetComponent(orphanComponentId, componentValue) Then
                orphanNames.Add candidateShape.Name
            End If
        End If
    Next candidateShape
    On Error Resume Next
    For Each orphanName In orphanNames
        mSheet.Shapes(CStr(orphanName)).Delete
    Next orphanName
    On Error GoTo 0
End Sub

Friend Sub NotifyDirty(ByVal targetComponentId As String)
    Dim existing As Variant

    If TryGetItem(mDirtyIds, targetComponentId, existing) Then Exit Sub
    mDirtyIds.Add targetComponentId, targetComponentId
End Sub

Friend Sub NotifyComponentTouched(ByVal targetComponentId As String)
    NotifyDirty targetComponentId
    If mBatchDepth = 0 And mRendered Then FlushDirty
End Sub

Public Sub FlushDirty()
    Dim componentValue As ReDimUI
    Dim dirtyId As Variant
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
    For Each dirtyId In mDirtyIds
        If TryGetComponent(CStr(dirtyId), componentValue) Then
            componentValue.Apply mSheet, mTheme, Me
        End If
    Next dirtyId
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
    ' A painted canvas belongs to the theme too: without this, a theme
    ' swap left dark controls floating on the old light background.
    If mCanvasPainted Then
        On Error Resume Next
        mSheet.Cells.Interior.Color = mTheme.CanvasColor
        On Error GoTo 0
    End If
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
Public Function WindowTitle(ByVal titleText As String) As ReDimUI
    RequireRole ROLE_APP, "WindowTitle"
    mWindowTitle = titleText
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
    Dim registeredIds As Collection
    Dim windowId As Variant
    Dim componentValue As ReDimUI
    Dim staleIds As Collection
    Dim staleId As Variant
    Dim known As Boolean
    Dim idx As Long
    Dim tabButton As ReDimUI

    If Not mHasNavBar Then Exit Sub
    Set registeredIds = ReDimUI.WindowIds
    BeginUpdate
    ' Tabs for windows that no longer exist get removed before rebuilding.
    Set staleIds = New Collection
    For Each componentValue In mComponents
        If Left$(componentValue.ComponentId, 4) = "nvb_" Then
            known = False
            For Each windowId In registeredIds
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
    For Each windowId In registeredIds
        idx = idx + 1
        Set tabButton = Button("nvb_" & CStr(windowId))
        tabButton.AtRect mNavBarLeft + (idx - 1) * (mNavBarTabWidth + 6), _
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
Public Function OnShow(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_APP, "OnShow"
    mOnShowProc = handlerProc
    Set OnShow = Me
End Function

Public Function OnHide(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_APP, "OnHide"
    mOnHideProc = handlerProc
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
' Locked canvas cells are also unselectable by default - no selection
' rectangle on the app surface, no protected-cell warnings for stray
' keys - while deliberately unlocked cells (cell-backed TextInput) stay
' selectable and editable; pass allowCellSelection:=True to keep the
' grid selectable. OnKey capture is unaffected either way. Opt-in, and
' scoped to this one sheet. Call after Render.
Public Function ProtectSurface( _
    Optional ByVal protectOn As Boolean = True, _
    Optional ByVal allowCellSelection As Boolean = False _
) As ReDimUI
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "ProtectSurface"
    If protectOn Then
        For Each componentValue In mComponents
            componentValue.UnlockInputCell mSheet
        Next componentValue
        mSheet.Protect DrawingObjects:=True, Contents:=True, _
            UserInterfaceOnly:=True
        If allowCellSelection Then
            mSheet.EnableSelection = xlNoRestrictions
        Else
            mSheet.EnableSelection = xlUnlockedCells
        End If
        mSurfaceProtected = True
    Else
        On Error Resume Next
        mSheet.Unprotect
        mSheet.EnableSelection = xlNoRestrictions
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
Public Sub HotKey(ByVal keyCode As String, ByVal handlerProc As String)
    RequireRole ROLE_APP, "HotKey"
    Application.OnKey keyCode, handlerProc
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
            componentValue.DeleteOwnShapes mSheet
        Next componentValue
    End If
    Set mComponents = New Collection
    Set mComponentOrder = New Collection
    Set mDirtyIds = New Collection
    ReDimUI.ForgetApp mAppId
End Sub

Friend Sub ForgetApp(ByVal targetAppId As String)
    Dim existing As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.ForgetApp targetAppId
        Exit Sub
    End If
    EnsureFactory
    If TryGetObject(mApps, targetAppId, existing) Then mApps.Remove targetAppId
    ForgetWindow targetAppId
End Sub

Friend Sub ReportAppError(ByVal messageText As String)
    RequireRole ROLE_APP, "ReportAppError"
    If LenB(mErrorHandlerProc) > 0 Then
        On Error Resume Next
        Application.Run mErrorHandlerProc, messageText
        On Error GoTo 0
    Else
        Application.StatusBar = messageText
    End If
End Sub

' Registers a one-argument workbook procedure that receives handler failures.
Public Sub OnError(ByVal handlerProc As String)
    RequireRole ROLE_APP, "OnError"
    mErrorHandlerProc = handlerProc
End Sub

' One app tick, split into isolated sections so a fault in one duty can
' never silently skip the rest: the compaction bug hid for months
' inside exactly that kind of swallowed abort. Faults stay silent to
' the user but count on the factory, where tests can prove there were
' none.
Friend Sub TickApp( _
    ByVal nowMs As LongLong, _
    ByVal elapsedMs As Long, _
    ByVal runWork As Boolean _
)
    TickOpsSection runWork
    TickJobsSection nowMs, runWork
    TickComponentsSection nowMs, elapsedMs, runWork
End Sub

Private Sub TickOpsSection(ByVal runWork As Boolean)
    Dim opValue As ReDimUI

    On Error GoTo Fault
    If Not runWork Then Exit Sub
    For Each opValue In mAsyncOps
        If opValue.OpIsRunning Then opValue.TickOp Me
    Next opValue
    Exit Sub
Fault:
    ReDimUI.NoteTickFault
End Sub

' Paced jobs self-gate on the wall clock, so sampling them every frame
' tightens their timing; budget jobs stay on the work cadence and give
' animations room while they run.
Private Sub TickJobsSection(ByVal nowMs As LongLong, ByVal runWork As Boolean)
    Dim jobValue As ReDimUI
    Dim componentValue As ReDimUI
    Dim animating As Boolean

    On Error GoTo Fault
    For Each componentValue In mComponents
        If componentValue.NeedsAnimation Then
            animating = True
            Exit For
        End If
    Next componentValue
    For Each jobValue In mJobs
        If jobValue.JobIsRunning Then
            jobValue.TickJob Me, nowMs, runWork, animating
        End If
    Next jobValue
    Exit Sub
Fault:
    ReDimUI.NoteTickFault
End Sub

Private Sub TickComponentsSection( _
    ByVal nowMs As LongLong, _
    ByVal elapsedMs As Long, _
    ByVal runWork As Boolean _
)
    Dim componentValue As ReDimUI
    Dim expiredIds As Collection
    Dim expiredId As Variant
    Dim leftButtonDown As Boolean
    Dim anyButtonDown As Boolean
    Dim pressEdge As Boolean

    On Error GoTo Fault
    ' Sliders drag with the left button only; the field blur watch treats
    ' a press of either button as clicking away.
    leftButtonDown = (GetAsyncKeyState(VK_LEFT_BUTTON) < 0)
    anyButtonDown = leftButtonDown _
        Or (GetAsyncKeyState(VK_RIGHT_BUTTON) < 0)
    pressEdge = (anyButtonDown And Not mFieldPrevButtonDown)
    mFieldPrevButtonDown = anyButtonDown
    If ReDimUI.ConsumeForcedPressEdge() Then pressEdge = True
    Set expiredIds = New Collection
    For Each componentValue In mComponents
        componentValue.TickAnimation mSheet, elapsedMs
        componentValue.TickSlideDrag Me, leftButtonDown
        componentValue.TickFieldWatch Me, anyButtonDown, pressEdge, runWork
        If runWork Then
            If componentValue.IsExpiredToast(nowMs) Then
                componentValue.BeginToastExit
            End If
        End If
        If componentValue.IsToastExitDone Then
            expiredIds.Add componentValue.ComponentId
        End If
    Next componentValue
    For Each expiredId In expiredIds
        If TryGetComponent(CStr(expiredId), componentValue) Then
            componentValue.Remove
        End If
    Next expiredId
    If expiredIds.Count > 0 Then CompactToastSlots
    Exit Sub
Fault:
    ReDimUI.NoteTickFault
End Sub

' After any toast leaves, survivors renumber to slots 1..N in their current
' order and slide up into the freed positions, the way a notification tray
' behaves.
Private Sub CompactToastSlots()
    Dim live As Collection
    Dim componentValue As ReDimUI
    Dim lowest As ReDimUI
    Dim idx As Long
    Dim lowestAt As Long
    Dim assigned As Long

    Set live = New Collection
    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then live.Add componentValue
    Next componentValue
    ' Collection.Item returns a Variant, and Friend members are invisible
    ' to late binding (error 438), so every candidate goes through a typed
    ' local before any member access. The failure only surfaces with two
    ' or more survivors, when the comparison loop actually runs.
    Dim candidate As ReDimUI
    Do While live.Count > 0
        Set lowest = live.Item(1)
        lowestAt = 1
        For idx = 2 To live.Count
            Set candidate = live.Item(idx)
            If candidate.ToastSlot < lowest.ToastSlot Then
                Set lowest = candidate
                lowestAt = idx
            End If
        Next idx
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

Private Function FindKeyed( _
    ByVal store As Collection, _
    ByVal lookupKey As String _
) As ReDimUI
    On Error Resume Next
    Set FindKeyed = store.Item(lookupKey)
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
    ByVal targetComponentId As String, _
    ByVal kindCode As Long _
)
    mRole = ROLE_COMPONENT
    mOwnerAppId = ownerAppId
    mComponentId = targetComponentId
    mKind = kindCode
    mVisible = True
    mEnabled = True
    mFillColor = -1
    mTextColor = -1
    mFontSize = 0
    mBusyText = "Working..."
    Select Case kindCode
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
            mListRows = DEFAULT_LIST_ROWS
        Case KIND_SELECT
            mWidth = 130: mHeight = 24
            mListRows = DEFAULT_LIST_ROWS
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
        Case KIND_TRANSFER
            mWidth = 360: mHeight = 140
        Case KIND_CHECKLIST
            mWidth = 170: mHeight = 100
            mSelectAllRow = True
        Case KIND_IMAGE
            mWidth = 120: mHeight = 80
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

Public Function At(ByVal rangeAddress As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "At"
    mAnchor = rangeAddress
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

Public Function Text(ByVal displayText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Text"
    mText = displayText
    Touch
    Set Text = Me
End Function

Public Function BusyText(ByVal displayText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BusyText"
    mBusyText = displayText
    Set BusyText = Me
End Function

Public Function FontSize(ByVal fontPoints As Double) As ReDimUI
    RequireRole ROLE_COMPONENT, "FontSize"
    mFontSize = fontPoints
    Touch
    Set FontSize = Me
End Function

Public Function Bold(Optional ByVal flag As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Bold"
    mFontBold = flag
    Touch
    Set Bold = Me
End Function

Public Function Fill(ByVal fillRgb As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "Fill"
    mFillColor = fillRgb
    mStyleVariant = VARIANT_NEUTRAL
    Touch
    Set Fill = Me
End Function

Public Function TextColor(ByVal inkRgb As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "TextColor"
    mTextColor = inkRgb
    Touch
    Set TextColor = Me
End Function

' The image's source file path. Loaded once per distinct path, embedded
' with the workbook, stretched to the declared rectangle.
Public Function Source(ByVal picturePath As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Source"
    mImageSource = picturePath
    Touch
    Set Source = Me
End Function

' Drives the image source from a state key, like the other bindings.
Public Function BindSource(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindSource"
    mBindSourceKey = stateKey
    Touch
    Set BindSource = Me
End Function

' Multi-line editing for a float TextInput: Enter inserts a newline,
' and committing moves to Tab, Ctrl+Enter, or clicking away. Esc still
' reverts. Size the rectangle for the lines you expect.
Public Function MultiLine(Optional ByVal multiLineOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "MultiLine"
    If mKind <> KIND_INPUT Then
        RaiseError ERROR_INVALID_OPERATION, "MultiLine", _
            "MultiLine applies to TextInput only."
    End If
    mMultiLine = multiLineOn
    Touch
    Set MultiLine = Me
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

Private Function SetVariant(ByVal variantCode As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "Variant"
    mStyleVariant = variantCode
    Touch
    Set SetVariant = Me
End Function

Public Function Visible(ByVal flag As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Visible"
    mVisible = flag
    Touch
    Set Visible = Me
End Function

Public Function Enabled(ByVal flag As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Enabled"
    mEnabled = flag
    Touch
    Set Enabled = Me
End Function

Public Function Value(ByVal newNumber As Double) As ReDimUI
    RequireRole ROLE_COMPONENT, "Value"
    mValue = newNumber
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
Public Function OnClick(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnClick"
    mClickProc = handlerProc
    Touch
    Set OnClick = Me
End Function

' Registers a zero-argument workbook procedure run through the async engine:
' the button disables itself and shows busy text while the work runs.
Public Function OnClickAsync(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnClickAsync"
    mAsyncClickProc = handlerProc
    Touch
    Set OnClickAsync = Me
End Function

' One-declaration navigation link: clicking this component navigates to
' the named window after any OnClick handler runs.
Public Function NavigatesTo(ByVal targetAppId As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "NavigatesTo"
    mNavTargetId = targetAppId
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

Public Function Checked(ByVal flag As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Checked"
    mChecked = flag
    Touch
    Set Checked = Me
End Function

Public Property Get IsChecked() As Boolean
    RequireRole ROLE_COMPONENT, "IsChecked"
    IsChecked = mChecked
End Property

Public Function Items(ParamArray listItems() As Variant) As ReDimUI
    Dim idx As Long

    RequireRole ROLE_COMPONENT, "Items"
    Set mItems = New Collection
    For idx = LBound(listItems) To UBound(listItems)
        mItems.Add CStr(listItems(idx))
    Next idx
    ClampItemSelection
    Touch
    Set Items = Me
End Function

' Replaces the item list from a 1D array, a Collection, a Range (one
' item per non-empty cell), or a ROneCOne sequence, so pickers can be
' fed straight from data.
Public Function ItemsFrom(ByVal itemSource As Variant) As ReDimUI
    RequireRole ROLE_COMPONENT, "ItemsFrom"
    Set mItems = ReadItemSource(itemSource, "ItemsFrom")
    ClampItemSelection
    If Not mCheckedSet Is Nothing Then Set mCheckedSet = New Collection
    Touch
    Set ItemsFrom = Me
End Function

' Fills the transfer list's chosen (right-hand) side from a 1D array, a
' Collection, a Range, or a ROneCOne sequence, mirroring ItemsFrom for
' the available side.
Public Function ChosenFrom(ByVal itemSource As Variant) As ReDimUI
    RequireRole ROLE_COMPONENT, "ChosenFrom"
    Set mChosenItems = ReadItemSource(itemSource, "ChosenFrom")
    Set mSelRight = New Collection
    Touch
    Set ChosenFrom = Me
End Function

' Seeds the checkbox list's checked rows by item text; names that match
' no item are ignored. Programmatic, so no state write and no OnChange.
Public Function CheckedFrom(ByVal itemSource As Variant) As ReDimUI
    Dim entry As Variant
    Dim idx As Long
    Dim wanted As Collection

    RequireRole ROLE_COMPONENT, "CheckedFrom"
    Set wanted = ReadItemSource(itemSource, "CheckedFrom")
    Set mCheckedSet = New Collection
    If Not mItems Is Nothing Then
        For idx = 1 To mItems.Count
            For Each entry In wanted
                If StrComp(CStr(mItems.Item(idx)), CStr(entry), _
                    vbTextCompare) = 0 Then
                    If Not TransferSelectionHas(mCheckedSet, idx) Then
                        mCheckedSet.Add idx
                    End If
                    Exit For
                End If
            Next entry
        Next idx
    End If
    Touch
    Set CheckedFrom = Me
End Function

' Reads an item source into a fresh Collection of strings. Accepts a 1D
' array, a Collection, a Range (one item per non-empty cell), or a
' ROneCOne sequence. Centralized so every picker reader shares one shape
' and the array test lives in exactly one guarded place.
Private Function ReadItemSource( _
    ByVal itemSource As Variant, _
    ByVal memberName As String _
) As Collection
    Dim gathered As Collection
    Dim entry As Variant
    Dim cell As Range
    Dim idx As Long

    Set gathered = New Collection
    If TypeName(itemSource) = "Range" Then
        For Each cell In itemSource.Cells
            If LenB(CStr(cell.Value)) > 0 Then gathered.Add CStr(cell.Value)
        Next cell
    ElseIf TypeName(itemSource) = "Collection" Then
        For Each entry In itemSource
            gathered.Add CStr(entry)
        Next entry
    ElseIf TypeName(itemSource) = "ROneCOne" Then
        ' A ROneCOne sequence enumerates like any other For Each source.
        For Each entry In itemSource
            gathered.Add CStr(entry)
        Next entry
    ElseIf IsArrayValue(itemSource) Then
        For idx = LBound(itemSource) To UBound(itemSource)
            gathered.Add CStr(itemSource(idx))
        Next idx
    Else
        RaiseError ERROR_INVALID_ARGUMENT, memberName, memberName & _
            " accepts an array, a Collection, a Range, or a ROneCOne sequence."
    End If
    Set ReadItemSource = gathered
End Function

' The one guarded home for the array test. An object is never an array,
' and asking IsArray about a ROneCOne dereferences its default member,
' which raises on the zero arguments the dereference supplies; the raise
' is then swallowed by any active On Error, leaving Err dirty for the
' next reader. The IsObject short-circuit keeps the object out of IsArray.
Private Function IsArrayValue(ByVal candidate As Variant) As Boolean
    If IsObject(candidate) Then Exit Function
    IsArrayValue = IsArray(candidate)
End Function

' Renders or hides the select-all header row (on by default).
Public Function WithSelectAll( _
    Optional ByVal selectAllOn As Boolean = True _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "WithSelectAll"
    mSelectAllRow = selectAllOn
    Touch
    Set WithSelectAll = Me
End Function

' Programmatic check state for one row: renders, but writes no state and
' fires no OnChange, like every programmatic item mutation.
Public Function SetItemChecked( _
    ByVal itemNumber As Long, _
    ByVal checkedFlag As Boolean _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "SetItemChecked"
    EnsureCheckedSet
    If itemNumber >= 1 And itemNumber <= ItemCount Then
        If checkedFlag Then
            If Not TransferSelectionHas(mCheckedSet, itemNumber) Then
                mCheckedSet.Add itemNumber
            End If
        Else
            If TransferSelectionHas(mCheckedSet, itemNumber) Then
                TransferSelectionToggle mCheckedSet, itemNumber
            End If
        End If
        Touch
    End If
    Set SetItemChecked = Me
End Function

Public Property Get IsItemChecked(ByVal itemNumber As Long) As Boolean
    RequireRole ROLE_COMPONENT, "IsItemChecked"
    If mCheckedSet Is Nothing Then Exit Property
    IsItemChecked = TransferSelectionHas(mCheckedSet, itemNumber)
End Property

Public Property Get CheckedCount() As Long
    RequireRole ROLE_COMPONENT, "CheckedCount"
    If Not mCheckedSet Is Nothing Then CheckedCount = mCheckedSet.Count
End Property

' Header captions for the transfer list panels; each header also shows the
' live item count.
Public Function Captions( _
    ByVal leftText As String, _
    ByVal rightText As String _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "Captions"
    mCaptionLeft = leftText
    mCaptionRight = rightText
    Touch
    Set Captions = Me
End Function

Public Property Get ChosenCount() As Long
    RequireRole ROLE_COMPONENT, "ChosenCount"
    If Not mChosenItems Is Nothing Then ChosenCount = mChosenItems.Count
End Property

Public Property Get ChosenTextAt(ByVal itemNumber As Long) As String
    RequireRole ROLE_COMPONENT, "ChosenTextAt"
    If mChosenItems Is Nothing Then Exit Property
    If itemNumber < 1 Or itemNumber > mChosenItems.Count Then Exit Property
    ChosenTextAt = CStr(mChosenItems.Item(itemNumber))
End Property

' Appends, or inserts at a 1-based position. The selected item stays
' selected, and checked rows stay checked: an insert at or before them
' shifts their indices along with them.
Public Function AddItem( _
    ByVal itemText As String, _
    Optional ByVal atPosition As Long = 0 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "AddItem"
    If mItems Is Nothing Then Set mItems = New Collection
    If atPosition >= 1 And atPosition <= mItems.Count Then
        mItems.Add itemText, Before:=atPosition
        If CLng(mValue) >= atPosition Then mValue = mValue + 1
        ShiftCheckedForInsert atPosition
    Else
        mItems.Add itemText
    End If
    Touch
    Set AddItem = Me
End Function

Private Sub ShiftCheckedForInsert(ByVal atPosition As Long)
    Dim idx As Long

    If mCheckedSet Is Nothing Then Exit Sub
    For idx = 1 To mCheckedSet.Count
        If CLng(mCheckedSet.Item(idx)) >= atPosition Then
            mCheckedSet.Add CLng(mCheckedSet.Item(idx)) + 1, _
                Before:=idx
            mCheckedSet.Remove idx + 1
        End If
    Next idx
End Sub

Private Sub ShiftCheckedForRemove(ByVal removedAt As Long)
    Dim idx As Long

    If mCheckedSet Is Nothing Then Exit Sub
    For idx = mCheckedSet.Count To 1 Step -1
        If CLng(mCheckedSet.Item(idx)) = removedAt Then
            mCheckedSet.Remove idx
        ElseIf CLng(mCheckedSet.Item(idx)) > removedAt Then
            mCheckedSet.Add CLng(mCheckedSet.Item(idx)) - 1, _
                Before:=idx
            mCheckedSet.Remove idx + 1
        End If
    Next idx
End Sub

' Removes by 1-based position, or by first matching text. Removing an item
' before the selection shifts the index; removing the selected item clears
' the selection to the placeholder rather than guessing a neighbor.
Public Function RemoveItem(ByVal indexOrText As Variant) As ReDimUI
    Dim idx As Long
    Dim probe As Long

    RequireRole ROLE_COMPONENT, "RemoveItem"
    If mItems Is Nothing Then
        Set RemoveItem = Me
        Exit Function
    End If
    If IsNumeric(indexOrText) Then
        idx = CLng(indexOrText)
    Else
        For probe = 1 To mItems.Count
            If CStr(mItems.Item(probe)) = CStr(indexOrText) Then
                idx = probe
                Exit For
            End If
        Next probe
    End If
    If idx < 1 Or idx > mItems.Count Then
        Set RemoveItem = Me
        Exit Function
    End If
    mItems.Remove idx
    If CLng(mValue) = idx Then
        mValue = 0
    ElseIf CLng(mValue) > idx Then
        mValue = mValue - 1
    End If
    ShiftCheckedForRemove idx
    Touch
    Set RemoveItem = Me
End Function

Public Function ClearItems() As ReDimUI
    RequireRole ROLE_COMPONENT, "ClearItems"
    Set mItems = New Collection
    mValue = 0
    If Not mCheckedSet Is Nothing Then Set mCheckedSet = New Collection
    Touch
    Set ClearItems = Me
End Function

Public Property Get ItemCount() As Long
    RequireRole ROLE_COMPONENT, "ItemCount"
    If Not mItems Is Nothing Then ItemCount = mItems.Count
End Property

Public Property Get ItemTextAt(ByVal itemNumber As Long) As String
    RequireRole ROLE_COMPONENT, "ItemTextAt"
    ItemTextAt = ItemAt(itemNumber)
End Property

Private Sub ClampItemSelection()
    If mItems Is Nothing Then
        mValue = 0
    ElseIf CLng(mValue) > mItems.Count Then
        mValue = 0
    End If
End Sub

' Item rows a SelectBox or ComboBox drop list shows at once, eight by
' default. A longer list windows: pager rows at the list edges count
' the items beyond them and page the window when clicked.
Public Function ListRows(ByVal shownRows As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "ListRows"
    If mKind <> KIND_SELECT And mKind <> KIND_COMBO Then
        RaiseError ERROR_INVALID_OPERATION, "ListRows", _
            "ListRows applies to SelectBox and ComboBox only."
    End If
    If shownRows < 1 Then
        RaiseError ERROR_INVALID_ARGUMENT, "ListRows", _
            "ListRows needs at least one row."
    End If
    mListRows = shownRows
    Touch
    Set ListRows = Me
End Function

Public Function SliderRange( _
    ByVal lowestValue As Long, _
    ByVal highestValue As Long, _
    Optional ByVal stepSize As Long = 1 _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "SliderRange"
    If highestValue <= lowestValue Then
        RaiseError ERROR_INVALID_ARGUMENT, "SliderRange", _
            "maximum must be greater than minimum."
    End If
    mSliderMin = lowestValue
    mSliderMax = highestValue
    mSliderStep = stepSize
    Touch
    Set SliderRange = Me
End Function

' OnChange is OnClick under a truthful name for value controls: form control
' OnAction fires after the value changes.
Public Function OnChange(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnChange"
    mClickProc = handlerProc
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

Public Property Let InputValue(ByVal assignedText As String)
    Dim priorEvents As Boolean

    RequireRole ROLE_COMPONENT, "InputValue"
    If IsFloatField() Then
        mText = assignedText
        Touch
    Else
        priorEvents = Application.EnableEvents
        Application.EnableEvents = False
        InputCell(OwnerApp.Sheet).Value = assignedText
        Application.EnableEvents = priorEvents
    End If
    If LenB(mWritesToKey) > 0 Then OwnerApp.SetState mWritesToKey, assignedText
End Property

Friend Function InputCell(ByVal hostSheet As Worksheet) As Range
    If LenB(mAnchor) = 0 Then
        RaiseError ERROR_INVALID_OPERATION, "TextInput", _
            "TextInput '" & mComponentId & "' needs At with a cell anchor."
    End If
    Set InputCell = hostSheet.Range(mAnchor).Cells(1, 1)
End Function

' A cell-backed field's shape sits over its anchor cell and takes the
' click, so a click on the field selects the cell itself. Only the active
' sheet can take a selection; programmatic dispatch elsewhere is a no-op.
Private Sub SelectInputCell(ByVal hostSheet As Worksheet)
    Dim cell As Range

    On Error Resume Next
    Set cell = InputCell(hostSheet)
    If Not cell Is Nothing Then
        If hostSheet Is ActiveSheet Then cell.Select
    End If
    On Error GoTo 0
End Sub

Friend Sub SyncInputFromSheet(ByVal hostApp As ReDimUI, ByVal changed As Range)
    Dim cell As Range
    Dim cellText As String
    Dim exactAt As Long

    If mKind <> KIND_INPUT And mKind <> KIND_COMBO Then Exit Sub
    If LenB(mAnchor) = 0 Then Exit Sub
    On Error Resume Next
    Set cell = InputCell(hostApp.Sheet)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub
    If Application.Intersect(cell, changed) Is Nothing Then Exit Sub
    cellText = CStr(cell.Value)
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, cellText
    If LenB(mClickProc) > 0 Then
        On Error Resume Next
        Application.Run mClickProc
        On Error GoTo 0
    End If
    ' Combo auto-suggest at commit time, the moment Excel hands control
    ' back: an exact match takes the item and closes; partial matches drop
    ' the list filtered by what was typed; no match stays closed as free
    ' text.
    If mKind = KIND_COMBO Then
        exactAt = ExactItemMatch(cellText)
        If exactAt > 0 Then
            mValue = exactAt
            mSelectOpen = False
        ElseIf LenB(cellText) > 0 _
            And FilterItems(cellText).Count > 0 Then
            mValue = 0
            mSelectOpen = True
            mFilterPristine = False
            mListOffset = 0
            mListHighlight = 0
        Else
            mValue = 0
            mSelectOpen = False
        End If
        Touch
    End If
End Sub

Friend Sub ConfigureToast( _
    ByVal messageText As String, _
    ByVal ttlMs As Long, _
    ByVal slot As Long _
)
    Dim hostApp As ReDimUI
    Dim extentLeft As Double
    Dim extentTop As Double
    Dim extentRight As Double
    Dim viewWidth As Double
    Dim viewHeight As Double

    mText = messageText
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
    ' An explicit ToastTray anchor is already a stable pin and always
    ' wins; the sticky rail applies only to the computed, clamped rail.
    If Not hostApp.ToastTrayOrigin(originLeft, originTop) Then
        If Not hostApp.ActiveToastRail(mComponentId, originLeft, originTop) Then
            If hostApp.ContentBounds(extentLeft, extentTop, extentRight) Then
                originLeft = extentRight + 12
                originTop = extentTop
            Else
                hostApp.ResolveViewport extentLeft, extentTop, viewWidth, _
                    viewHeight
                originLeft = extentLeft + viewWidth - mWidth - 12
                originTop = extentTop + 8
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
        hostApp.RecordToastRail originLeft, originTop
    End If
    mLeft = originLeft
    mTop = originTop + (slot - 1) * (mHeight + 6)
    ' Entrance animation: the shape spawns below its slot and the pump
    ' eases it up into place. During a compaction glide the newcomer's
    ' slot may still be visually occupied by a survivor easing upward, so
    ' the entrance starts below the column as drawn, not just the model:
    ' one full pitch under the lowest live toast when that is deeper than
    ' the default hop.
    Dim occupiedTop As Double
    mSlideRemaining = 14
    mToastExiting = False
    mExitRemaining = 0
    occupiedTop = hostApp.MaxToastVisualTop(mComponentId)
    If occupiedTop >= 0 Then
        If occupiedTop + mHeight + 6 - mTop > mSlideRemaining Then
            mSlideRemaining = occupiedTop + mHeight + 6 - mTop
        End If
    End If
    mHasRect = True
    mVisible = True
End Sub

Friend Property Get ToastSlot() As Long
    If mKind = KIND_TOAST Then ToastSlot = CLng(mValue)
End Property

Friend Property Get ToastVisualTop() As Double
    If mKind = KIND_TOAST Then ToastVisualTop = mTop + mSlideRemaining
End Property

Friend Property Get IsToastExiting() As Boolean
    IsToastExiting = (mKind = KIND_TOAST And mToastExiting)
End Property

' An expired or dismissed toast fades for a beat instead of popping out
' of existence; removal and compaction follow when the fade completes.
Friend Sub BeginToastExit()
    If mKind <> KIND_TOAST Then Exit Sub
    If mToastExiting Then Exit Sub
    mToastExiting = True
    mExitRemaining = 1
    ReDimUI.RequestPump
End Sub

Friend Property Get IsToastExitDone() As Boolean
    IsToastExitDone = (mKind = KIND_TOAST And mToastExiting _
        And mExitRemaining <= 0)
End Property

' Moves this toast to a new slot, easing from wherever it currently sits
' (including mid-entrance) to the new position without a visual jump.
Friend Sub ReslotToast(ByVal newSlot As Long)
    Dim slotPitch As Double
    Dim originTop As Double
    Dim visualTop As Double

    If mKind <> KIND_TOAST Then Exit Sub
    slotPitch = mHeight + 6
    originTop = mTop - (CLng(mValue) - 1) * slotPitch
    visualTop = mTop + mSlideRemaining
    mValue = newSlot
    mTop = originTop + (newSlot - 1) * slotPitch
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
    DeleteOwnShapes hostApp.Sheet
    hostApp.ForgetComponent mComponentId
End Sub

Friend Sub ForgetComponent(ByVal targetComponentId As String)
    Dim idx As Long

    RequireRole ROLE_APP, "ForgetComponent"
    On Error Resume Next
    mComponents.Remove targetComponentId
    On Error GoTo 0
    For idx = mComponentOrder.Count To 1 Step -1
        If mComponentOrder.Item(idx) = targetComponentId Then
            mComponentOrder.Remove idx
        End If
    Next idx
End Sub

Friend Function IsBoundTo(ByVal stateKey As String) As Boolean
    IsBoundTo = (mBindTextKey = stateKey) Or (mBindValueKey = stateKey) _
        Or (mBindVisibleKey = stateKey) Or (mBindEnabledKey = stateKey) _
        Or (LenB(mBindSourceKey) > 0 And mBindSourceKey = stateKey)
End Function

' =====================================================================
' Transfer list moves
' =====================================================================

Private Sub EnsureTransferCollections()
    If mItems Is Nothing Then Set mItems = New Collection
    If mChosenItems Is Nothing Then Set mChosenItems = New Collection
    If mSelLeft Is Nothing Then Set mSelLeft = New Collection
    If mSelRight Is Nothing Then Set mSelRight = New Collection
End Sub

' True when floating chrome above the canvas claims the point, asked by
' the pump watches before a geometric hit test engages a control that
' the chrome may be covering.
Friend Function PointClaimedByChrome( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double, _
    ByVal exceptComponentId As String _
) As Boolean
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "PointClaimedByChrome"
    For Each componentValue In mComponents
        If componentValue.ComponentId <> exceptComponentId Then
            If componentValue.ChromeClaimsPoint(pointsX, pointsY) Then
                PointClaimedByChrome = True
                Exit Function
            End If
        End If
    Next componentValue
End Function

' Rows a transfer panel can draw; mirrors the renderer's header height
' (18), row height (20), and row inset (6). Keep in sync with
' ApplyTransferParts and SyncTransferRows.
Private Function TransferRowCapacity() As Long
    Dim panelH As Double

    panelH = mHeight - 18
    If panelH < 30 Then panelH = 30
    TransferRowCapacity = Int((panelH - 6) / 20)
    If TransferRowCapacity < 1 Then TransferRowCapacity = 1
End Function

Private Sub EnsureCheckedSet()
    If mCheckedSet Is Nothing Then Set mCheckedSet = New Collection
End Sub

' Collection values joined into a repaint signature; values only, so no
' late-bound member access is involved.
Private Function JoinedSig(ByVal store As Collection) As String
    Dim idx As Long

    For idx = 1 To store.Count
        JoinedSig = JoinedSig & "|" & CStr(store.Item(idx))
    Next idx
End Function

' Checked item texts in list order, joined for the WritesTo state value.
Private Function CheckedJoined() As String
    Dim idx As Long

    If mItems Is Nothing Then Exit Function
    For idx = 1 To mItems.Count
        If TransferSelectionHas(mCheckedSet, idx) Then
            If LenB(CheckedJoined) > 0 Then
                CheckedJoined = CheckedJoined & ", "
            End If
            CheckedJoined = CheckedJoined & CStr(mItems.Item(idx))
        End If
    Next idx
End Function

Private Function TransferSelectionHas( _
    ByVal selectionSet As Collection, _
    ByVal rowNumber As Long _
) As Boolean
    Dim entry As Variant

    For Each entry In selectionSet
        If CLng(entry) = rowNumber Then
            TransferSelectionHas = True
            Exit Function
        End If
    Next entry
End Function

Private Sub TransferSelectionToggle( _
    ByVal selectionSet As Collection, _
    ByVal rowNumber As Long _
)
    Dim idx As Long

    For idx = 1 To selectionSet.Count
        If CLng(selectionSet.Item(idx)) = rowNumber Then
            selectionSet.Remove idx
            Exit Sub
        End If
    Next idx
    selectionSet.Add rowNumber
End Sub

Private Sub PruneTransferSelection( _
    ByVal selectionSet As Collection, _
    ByVal maxIndex As Long _
)
    Dim idx As Long

    For idx = selectionSet.Count To 1 Step -1
        If CLng(selectionSet.Item(idx)) > maxIndex Then
            selectionSet.Remove idx
        End If
    Next idx
End Sub

' Moves every selected row. The ascending pass keeps the rows' relative
' order in the target; the descending removal pass keeps yet-unremoved
' source indices stable.
Private Sub TransferSelected(ByVal hostApp As ReDimUI, ByVal rightward As Boolean)
    Dim sourceItems As Collection
    Dim targetItems As Collection
    Dim selectionSet As Collection
    Dim ordered As Collection
    Dim idx As Long

    If rightward Then
        Set sourceItems = mItems
        Set targetItems = mChosenItems
        Set selectionSet = mSelLeft
    Else
        Set sourceItems = mChosenItems
        Set targetItems = mItems
        Set selectionSet = mSelRight
    End If
    Set ordered = New Collection
    For idx = 1 To sourceItems.Count
        If TransferSelectionHas(selectionSet, idx) Then
            ordered.Add CStr(sourceItems.Item(idx))
        End If
    Next idx
    If ordered.Count = 0 Then Exit Sub
    For idx = sourceItems.Count To 1 Step -1
        If TransferSelectionHas(selectionSet, idx) Then
            sourceItems.Remove idx
        End If
    Next idx
    For idx = 1 To ordered.Count
        targetItems.Add CStr(ordered.Item(idx))
    Next idx
    If rightward Then
        Set mSelLeft = New Collection
    Else
        Set mSelRight = New Collection
    End If
    Touch
    CommitTransferChange hostApp
End Sub

Private Sub TransferAll(ByVal hostApp As ReDimUI, ByVal rightward As Boolean)
    Dim idx As Long

    If rightward Then
        If mItems.Count = 0 Then Exit Sub
        For idx = 1 To mItems.Count
            mChosenItems.Add CStr(mItems.Item(idx))
        Next idx
        Set mItems = New Collection
    Else
        If mChosenItems.Count = 0 Then Exit Sub
        For idx = 1 To mChosenItems.Count
            mItems.Add CStr(mChosenItems.Item(idx))
        Next idx
        Set mChosenItems = New Collection
    End If
    Set mSelLeft = New Collection
    Set mSelRight = New Collection
    Touch
    CommitTransferChange hostApp
End Sub

' A user transfer writes the chosen list (joined with a comma and space)
' to state and fires OnChange once. Selection clicks fire nothing.
Private Sub CommitTransferChange(ByVal hostApp As ReDimUI)
    Dim joined As String
    Dim idx As Long

    For idx = 1 To mChosenItems.Count
        If idx > 1 Then joined = joined & ", "
        joined = joined & CStr(mChosenItems.Item(idx))
    Next idx
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, joined
    If LenB(mClickProc) > 0 Then
        On Error Resume Next
        Application.Run mClickProc
        On Error GoTo 0
    End If
End Sub

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
            Else
                ' The frame covers the anchor cell and takes the click, so
                ' the click selects that cell: typing lands in the grid and
                ' F2 edits in place, as if the cell itself was clicked.
                SelectInputCell hostApp.Sheet
            End If
            Exit Sub
        Case KIND_COMBO
            If ReDimUI.SenderPart = "optu" Or ReDimUI.SenderPart = "optd" Then
                PageDropList (ReDimUI.SenderPart = "optd")
                Exit Sub
            ElseIf Left$(ReDimUI.SenderPart, 3) = "opt" _
                And IsNumeric(Mid$(ReDimUI.SenderPart, 4)) Then
                ' Clicked window rows map through the scroll offset to
                ' their absolute match index. A pick that leaves the value
                ' as it was runs no handlers.
                If Not PickComboItem(hostApp, _
                    mListOffset + CLng(Mid$(ReDimUI.SenderPart, 4))) Then
                    Exit Sub
                End If
            ElseIf IsFloatField() Then
                If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
                    ToggleDropList hostApp.Sheet
                Else
                    FocusField hostApp
                End If
                Exit Sub
            Else
                ' Cell-backed: the face covers the anchor cell, so a click
                ' selects the cell for typing as well as toggling the list.
                SelectInputCell hostApp.Sheet
                ToggleDropList hostApp.Sheet
                Exit Sub
            End If
        Case KIND_SELECT
            If ReDimUI.SenderPart = "optu" Or ReDimUI.SenderPart = "optd" Then
                PageDropList (ReDimUI.SenderPart = "optd")
                Exit Sub
            ElseIf Left$(ReDimUI.SenderPart, 3) = "opt" _
                And IsNumeric(Mid$(ReDimUI.SenderPart, 4)) Then
                Dim pickedItem As Long
                pickedItem = ListItemAtRow(CLng(Mid$(ReDimUI.SenderPart, 4)))
                If pickedItem = 0 Then Exit Sub
                ' Re-picking the selection only closes the list: nothing
                ' changed, so nothing is written and no handler runs.
                If pickedItem = CLng(mValue) Then
                    mSelectOpen = False
                    Touch
                    Exit Sub
                End If
                mValue = pickedItem
                mSelectOpen = False
                If LenB(mWritesToKey) > 0 Then
                    hostApp.SetState mWritesToKey, ItemAt(pickedItem)
                End If
                Touch
            Else
                ToggleDropList hostApp.Sheet
                Exit Sub
            End If
        Case KIND_TICKBOX
            mChecked = Not mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
            Touch
        Case KIND_CHECKLIST
            ' Select-all follows the convention: mixed or none checks
            ' everything, only the all-checked state unchecks. Falls
            ' through so OnChange fires once per toggle, bulk included.
            Dim checkPart As String
            Dim checkRow As Long
            checkPart = ReDimUI.SenderPart
            EnsureCheckedSet
            If checkPart = "mb" Or checkPart = "mt" Then
                If ItemCount = 0 Then Exit Sub
                If mCheckedSet.Count = ItemCount Then
                    Set mCheckedSet = New Collection
                Else
                    Set mCheckedSet = New Collection
                    For checkRow = 1 To ItemCount
                        mCheckedSet.Add checkRow
                    Next checkRow
                End If
            Else
                If LenB(checkPart) = 0 Then
                    checkRow = 1
                ElseIf IsNumeric(Mid$(checkPart, 2)) Then
                    checkRow = CLng(Mid$(checkPart, 2))
                End If
                If checkRow < 1 Or checkRow > ItemCount Then Exit Sub
                TransferSelectionToggle mCheckedSet, checkRow
            End If
            If LenB(mWritesToKey) > 0 Then
                hostApp.SetState mWritesToKey, CheckedJoined()
            End If
            Touch
        Case KIND_TRANSFER
            Dim transferPart As String
            Dim transferRow As Long
            transferPart = ReDimUI.SenderPart
            EnsureTransferCollections
            Dim transferPage As Long
            transferPage = TransferRowCapacity() - 1
            If transferPage < 1 Then transferPage = 1
            If transferPart = "alu" Then
                mLeftOffset = mLeftOffset - transferPage
                If mLeftOffset < 0 Then mLeftOffset = 0
                Touch
            ElseIf transferPart = "ald" Then
                mLeftOffset = mLeftOffset + transferPage
                Touch
            ElseIf transferPart = "clu" Then
                mRightOffset = mRightOffset - transferPage
                If mRightOffset < 0 Then mRightOffset = 0
                Touch
            ElseIf transferPart = "cld" Then
                mRightOffset = mRightOffset + transferPage
                Touch
            ElseIf Left$(transferPart, 2) = "al" Then
                If IsNumeric(Mid$(transferPart, 3)) Then
                    ' Window rows map through the scroll offset.
                    transferRow = mLeftOffset + CLng(Mid$(transferPart, 3))
                    TransferSelectionToggle mSelLeft, transferRow
                    Touch
                End If
            ElseIf Left$(transferPart, 2) = "cl" Then
                If IsNumeric(Mid$(transferPart, 3)) Then
                    transferRow = mRightOffset + CLng(Mid$(transferPart, 3))
                    TransferSelectionToggle mSelRight, transferRow
                    Touch
                End If
            ElseIf transferPart = "mvr" Then
                TransferSelected hostApp, True
            ElseIf transferPart = "mvl" Then
                TransferSelected hostApp, False
            ElseIf transferPart = "mvar" Then
                TransferAll hostApp, True
            ElseIf transferPart = "mval" Then
                TransferAll hostApp, False
            End If
            Exit Sub
        Case KIND_RADIO
            Dim rowNumber As Long
            If LenB(ReDimUI.SenderPart) = 0 Then
                rowNumber = 1
            ElseIf IsNumeric(Mid$(ReDimUI.SenderPart, 2)) Then
                rowNumber = CLng(Mid$(ReDimUI.SenderPart, 2))
            End If
            If rowNumber < 1 Then Exit Sub
            If rowNumber = CLng(mValue) Then Exit Sub
            mValue = rowNumber
            If LenB(mWritesToKey) > 0 Then
                hostApp.SetState mWritesToKey, ItemAt(rowNumber)
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
    RunClickHandlers hostApp
End Sub

' What a click, or a keyboard pick standing in for one, runs after the
' component's own work: the async op when OnClickAsync registered one,
' otherwise the OnClick or OnChange procedure, then any modal close and
' navigation.
Private Sub RunClickHandlers(ByVal hostApp As ReDimUI)
    If LenB(mAsyncClickProc) > 0 Then
        hostApp.StartAsyncClick Me, mAsyncClickProc
        Exit Sub
    End If
    If LenB(mClickProc) > 0 Then Application.Run mClickProc
    If mCloseModalOnClick Then hostApp.CloseModal
    If LenB(mNavTargetId) > 0 Then ReDimUI.Navigate mNavTargetId
End Sub

Private Function ItemAt(ByVal itemNumber As Long) As String
    If mItems Is Nothing Then Exit Function
    If itemNumber < 1 Or itemNumber > mItems.Count Then Exit Function
    ItemAt = CStr(mItems.Item(itemNumber))
End Function

' The OnClickAsync path: a per-button op that disables the button, runs the
' procedure through the pump, and restores the button on completion.
Friend Sub StartAsyncClick( _
    ByVal clickedComponent As ReDimUI, _
    ByVal handlerProc As String _
)
    Dim opValue As ReDimUI

    RequireRole ROLE_APP, "StartAsyncClick"
    Set opValue = Async("click_" & clickedComponent.ComponentId)
    If opValue.OpIsRunning Then Exit Sub
    opValue.ConfigureClickOp handlerProc, clickedComponent.ComponentId
    opValue.Start
End Sub

Friend Sub SetBusy(ByVal flag As Boolean)
    mBusy = flag
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
    ByVal hostApp As ReDimUI _
)
    Dim mainShape As Shape

    RequireRole ROLE_COMPONENT, "Apply"
    ResolveBindings hostApp
    EnsureGeometryResolved hostSheet
    Set mainShape = EnsureShape(hostSheet)
    ApplyGeometry mainShape
    ApplyStyle mainShape, themeValue
    ApplyText mainShape, themeValue
    ApplyBehavior mainShape, hostSheet
    ApplyKindExtras mainShape, hostSheet, themeValue
    mAppliedValid = True
End Sub

Private Sub ResolveBindings(ByVal hostApp As ReDimUI)
    Dim bound As Variant

    If LenB(mBindTextKey) > 0 Then
        If hostApp.HasState(mBindTextKey) Then
            mText = FormatBoundText(hostApp.State(mBindTextKey))
        End If
    End If
    If LenB(mBindValueKey) > 0 Then
        If hostApp.HasState(mBindValueKey) Then
            bound = hostApp.State(mBindValueKey)
            If IsNumeric(bound) Then mValue = CDbl(bound)
        End If
    End If
    If LenB(mBindVisibleKey) > 0 Then
        If hostApp.HasState(mBindVisibleKey) Then
            mVisible = CBool(hostApp.State(mBindVisibleKey)) Xor mBindVisibleInvert
        End If
    End If
    If LenB(mBindSourceKey) > 0 Then
        If hostApp.HasState(mBindSourceKey) Then
            mImageSource = CStr(hostApp.State(mBindSourceKey))
        End If
    End If
    If LenB(mBindEnabledKey) > 0 Then
        If hostApp.HasState(mBindEnabledKey) Then
            mEnabled = CBool(hostApp.State(mBindEnabledKey)) Xor mBindEnabledInvert
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
    Dim mainShape As Shape

    On Error Resume Next
    Set mainShape = hostSheet.Shapes(ShapeName)
    On Error GoTo 0
    If Not mainShape Is Nothing Then
        ' An adopted shape from an earlier framework version can carry the
        ' wrong geometry type; rebuild it rather than styling the wrong form.
        If mKind = KIND_SPINNER Then
            If mainShape.AutoShapeType <> msoShapeBlockArc Then
                mainShape.Delete
                Set mainShape = Nothing
            End If
        End If
    End If
    If mainShape Is Nothing Then
        Set mainShape = CreateShapeForKind(hostSheet)
        mainShape.Name = ShapeName
        mAppliedValid = False
    End If
    Set EnsureShape = mainShape
End Function

Private Function CreateShapeForKind(ByVal hostSheet As Worksheet) As Shape
    Dim mainShape As Shape

    Select Case mKind
        Case KIND_BUTTON
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.25
        Case KIND_LABEL
            Set mainShape = hostSheet.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, mLeft, mTop, mWidth, mHeight)
        Case KIND_CARD
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.08
        Case KIND_SPINNER
            ' A block arc is a filled ring segment concentric with its
            ' bounding box, so rotation stays visually circular. A plain
            ' arc's path center is off the box center and wobbles.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeBlockArc, mLeft, mTop, mWidth, mHeight)
            On Error Resume Next
            mainShape.Adjustments(1) = 0
            mainShape.Adjustments(2) = 290
            mainShape.Adjustments(3) = 0.34
            On Error GoTo 0
        Case KIND_TOGGLE
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.5
        Case KIND_TOAST
            ' Born at the entrance position, so the first paint cannot
            ' flash the shape at its settled slot before the offset lands.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop + mSlideRemaining, _
                mWidth, mHeight)
            mainShape.Adjustments(1) = 0.2
        Case KIND_SELECT
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.15
        Case KIND_TICKBOX
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, 14, 14)
            mainShape.Adjustments(1) = 0.2
        Case KIND_RADIO
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeOval, mLeft, mTop, 14, 14)
        Case KIND_STEPPER
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.15
        Case KIND_SLIDEBAR
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, 6)
            mainShape.Adjustments(1) = 0.5
        Case KIND_INPUT, KIND_COMBO
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.15
        Case KIND_TRANSFER
            ' The main shape is the left panel; extras own its geometry.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, 100, 100)
            mainShape.Adjustments(1) = 0.08
        Case KIND_CHECKLIST
            ' The main shape is the first row's box; extras own geometry.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, 14, 14)
            mainShape.Adjustments(1) = 0.2
        Case KIND_IMAGE
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.08
        Case Else
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRectangle, mLeft, mTop, mWidth, mHeight)
    End Select
    mainShape.Placement = xlFreeFloating
    Set CreateShapeForKind = mainShape
End Function

Private Function EnsurePartShape( _
    ByVal hostSheet As Worksheet, _
    ByVal partName As String, _
    ByVal shapeType As MsoAutoShapeType _
) As Shape
    Dim partShape As Shape
    Dim partShapeName As String

    partShapeName = ShapeName & PART_SEPARATOR & partName
    On Error Resume Next
    Set partShape = hostSheet.Shapes(partShapeName)
    On Error GoTo 0
    If partShape Is Nothing Then
        Set partShape = hostSheet.Shapes.AddShape(shapeType, mLeft, mTop, 8, 8)
        partShape.Name = partShapeName
        partShape.Placement = xlFreeFloating
        partShape.Line.Visible = msoFalse
    End If
    Set EnsurePartShape = partShape
End Function

Private Function IsCompositeDrawnKind() As Boolean
    IsCompositeDrawnKind = (mKind = KIND_TICKBOX Or mKind = KIND_RADIO _
        Or mKind = KIND_STEPPER Or mKind = KIND_SLIDEBAR _
        Or mKind = KIND_TRANSFER Or mKind = KIND_CHECKLIST)
End Function

' Geometry diffs against the shape's actual position rather than a cache, so
' a shape moved by hand (Ctrl+drag) snaps back to its declared rectangle on
' the next render. Composite drawn kinds own their geometry in extras: the
' main shape is one piece of several, not the declared rectangle.
Private Sub ApplyGeometry(ByVal mainShape As Shape)
    If IsCompositeDrawnKind() Then Exit Sub
    If Abs(mainShape.Left - mLeft) > 0.01 Then mainShape.Left = mLeft
    If Abs(mainShape.Top - mTop) > 0.01 Then mainShape.Top = mTop
    If Abs(mainShape.Width - mWidth) > 0.01 Then mainShape.Width = mWidth
    If Abs(mainShape.Height - mHeight) > 0.01 Then mainShape.Height = mHeight
    If Not mAppliedValid Or mAppliedVisible <> mVisible Then
        mainShape.Visible = BoolToMso(mVisible)
        mAppliedVisible = mVisible
    End If
End Sub

Private Sub ApplyStyle(ByVal mainShape As Shape, ByVal themeValue As ReDimUI)
    Dim resolvedFill As Long

    If IsCompositeDrawnKind() Then Exit Sub
    resolvedFill = ResolveFill(themeValue)
    Select Case mKind
        Case KIND_LABEL
            If Not mAppliedValid Then
                mainShape.Fill.Visible = msoFalse
                mainShape.Line.Visible = msoFalse
            End If
        Case KIND_SPINNER
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = ResolveSpinnerColor(themeValue)
                mainShape.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
        Case KIND_CARD, KIND_TOAST
            If Not mAppliedValid Or mAppliedFill <> ResolveSurfaceFill(themeValue) Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                mainShape.Line.Visible = msoTrue
                mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                mainShape.Line.Weight = 1
                mAppliedFill = ResolveSurfaceFill(themeValue)
            End If
        Case KIND_PROGRESS
            If Not mAppliedValid Or mAppliedFill <> themeValue.MutedColor Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = themeValue.MutedColor
                mainShape.Line.Visible = msoFalse
                mAppliedFill = themeValue.MutedColor
            End If
        Case KIND_TOGGLE
            If Not mAppliedValid Or mAppliedFill <> ResolveToggleFill(themeValue) Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = ResolveToggleFill(themeValue)
                mainShape.Line.Visible = msoFalse
                mAppliedFill = ResolveToggleFill(themeValue)
            End If
        Case KIND_IMAGE
            ApplyImageFill mainShape, themeValue
        Case KIND_INPUT, KIND_COMBO
            ' Shapes adopted from builds older than 0.9.1 are square;
            ' coerce them to the rounded profile the family shares.
            If Not mAppliedValid Then
                If mainShape.AutoShapeType <> msoShapeRoundedRectangle Then
                    mainShape.AutoShapeType = msoShapeRoundedRectangle
                    mainShape.Adjustments(1) = 0.15
                End If
            End If
            If IsFloatField() Then
                ' Float fields are their own editor: surface fill so the
                ' face takes clicks, accent border while focused. Diffed
                ' so a keystroke repaint costs no style writes at all.
                Dim ringOn As Boolean
                ringOn = ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId)
                If Not mAppliedValid _
                    Or mAppliedFill <> ResolveSurfaceFill(themeValue) _
                    Or mAppliedFocusRing <> ringOn Then
                    mainShape.Fill.Visible = msoTrue
                    mainShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                    mainShape.Line.Visible = msoTrue
                    If ringOn Then
                        mainShape.Line.ForeColor.RGB = themeValue.PrimaryColor
                        mainShape.Line.Weight = 1.5
                    Else
                        mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                        mainShape.Line.Weight = 1
                    End If
                    mAppliedFill = ResolveSurfaceFill(themeValue)
                    mAppliedFocusRing = ringOn
                End If
            ElseIf Not mAppliedValid Or mAppliedFill <> themeValue.BorderColor Then
                mainShape.Fill.Visible = msoFalse
                mainShape.Line.Visible = msoTrue
                mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                mainShape.Line.Weight = 1
                mAppliedFill = themeValue.BorderColor
            End If
        Case KIND_SELECT
            If Not mAppliedValid Or mAppliedFill <> ResolveSurfaceFill(themeValue) Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                mainShape.Line.Visible = msoTrue
                mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                mainShape.Line.Weight = 1
                mAppliedFill = ResolveSurfaceFill(themeValue)
            End If
        Case KIND_OVERLAY
            If Not mAppliedValid Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
                mainShape.Fill.Transparency = 0.45
                mainShape.Line.Visible = msoFalse
            End If
        Case Else
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = resolvedFill
                ' Surface-filled buttons need an outline or they vanish
                ' against cards and the canvas.
                If mStyleVariant = VARIANT_SECONDARY And mEnabled And Not mBusy Then
                    mainShape.Line.Visible = msoTrue
                    mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                    mainShape.Line.Weight = 1
                Else
                    mainShape.Line.Visible = msoFalse
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

' Paints the picture fill, loading the file only when the source path
' changed. A missing or unloadable source renders a themed placeholder,
' except when the shape already carries an embedded picture: the picture
' saved with the workbook outlives its source file by design.
Private Sub ApplyImageFill(ByVal mainShape As Shape, ByVal themeValue As ReDimUI)
    Dim resolvedSource As String
    Dim loaded As Boolean

    resolvedSource = mImageSource
    If LenB(resolvedSource) > 0 Then
        If FileExistsSafe(resolvedSource) Then
            If mAppliedImageSource = resolvedSource Then Exit Sub
            On Error Resume Next
            Err.Clear
            mainShape.Fill.Visible = msoTrue
            mainShape.Fill.UserPicture resolvedSource
            loaded = (Err.Number = 0)
            On Error GoTo 0
            If loaded Then
                mainShape.Line.Visible = msoFalse
                StyleGlyphText mainShape, vbNullString, _
                    themeValue.OnMutedColor, themeValue.BaseFontSize, _
                    themeValue
                mAppliedImageSource = resolvedSource
                Exit Sub
            End If
        End If
    End If
    If mainShape.Fill.Type = msoFillPicture Then
        mAppliedImageSource = IMAGE_MISSING
        Exit Sub
    End If
    If mAppliedImageSource = IMAGE_MISSING And mAppliedValid Then Exit Sub
    mainShape.Fill.Visible = msoTrue
    mainShape.Fill.Solid
    mainShape.Fill.ForeColor.RGB = themeValue.MutedColor
    mainShape.Line.Visible = msoTrue
    mainShape.Line.ForeColor.RGB = themeValue.BorderColor
    mainShape.Line.Weight = 1
    StyleGlyphText mainShape, "(image not found)", themeValue.OnMutedColor, _
        themeValue.BaseFontSize - 2, themeValue
    mAppliedImageSource = IMAGE_MISSING
End Sub

Private Function FileExistsSafe(ByVal candidatePath As String) As Boolean
    On Error Resume Next
    FileExistsSafe = (LenB(Dir$(candidatePath)) > 0)
    On Error GoTo 0
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

Friend Function ResolveText(ByVal hostApp As ReDimUI) As String
    If mBusy And mKind = KIND_BUTTON And LenB(mBusyText) > 0 Then
        ResolveText = mBusyText
        Exit Function
    End If
    If mKind = KIND_SELECT Then
        ResolveText = ItemAt(CLng(mValue))
        If LenB(ResolveText) = 0 Then ResolveText = mText
        Exit Function
    End If
    ' A focused float field shows the blinking insertion bar at the
    ' caret index; the blink swaps to a space so text never shifts.
    If IsFloatField() Then
        If ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
            If mCaretPos > Len(mText) Then mCaretPos = Len(mText)
            ResolveText = Left$(mText, mCaretPos) & _
                IIf(mCaretOn, "|", " ") & Mid$(mText, mCaretPos + 1)
            Exit Function
        End If
    End If
    ResolveText = mText
End Function

' The viewport of the buffer that fits the field, always containing the
' caret: the caret's hard line renders with a horizontal window around
' the caret, neighboring lines fill the remaining rows (truncated with
' an ellipsis when long, so line estimates stay exact), and ellipses at
' the edges mark trimmed content above and below. Character columns are
' estimated at 0.58 of the font size, lines at 1.35, conservatively.
Private Function WindowFieldText( _
    ByVal shownText As String, _
    ByVal resolvedSize As Double _
) As String
    Dim charsPerLine As Long
    Dim maxLines As Long
    Dim caretRender As Long
    Dim hardLines() As String
    Dim caretLine As Long
    Dim caretCol As Long
    Dim probe As Long
    Dim lineStart As Long
    Dim startLine As Long
    Dim endLine As Long
    Dim segment As String
    Dim viewText As String
    Dim idx As Long

    charsPerLine = Int((mWidth - 12) / (resolvedSize * 0.58))
    If charsPerLine < 4 Then charsPerLine = 4
    caretRender = mCaretPos + 1
    If Not mMultiLine Then
        WindowFieldText = WindowLineAroundCaret( _
            shownText, caretRender, charsPerLine)
        Exit Function
    End If
    maxLines = Int((mHeight - 6) / (resolvedSize * 1.35))
    If maxLines < 1 Then maxLines = 1
    hardLines = Split(shownText, vbLf)
    caretLine = 0
    lineStart = 0
    For probe = 1 To caretRender - 1
        If Mid$(shownText, probe, 1) = vbLf Then
            caretLine = caretLine + 1
            lineStart = probe
        End If
    Next probe
    caretCol = caretRender - lineStart
    startLine = caretLine - maxLines + 1
    If startLine < 0 Then startLine = 0
    endLine = startLine + maxLines - 1
    If endLine > UBound(hardLines) Then endLine = UBound(hardLines)
    For idx = startLine To endLine
        If idx = caretLine Then
            segment = WindowLineAroundCaret( _
                hardLines(idx), caretCol, charsPerLine)
        ElseIf Len(hardLines(idx)) > charsPerLine Then
            segment = Left$(hardLines(idx), charsPerLine - 1) & _
                ChrW(8230)
        Else
            segment = hardLines(idx)
        End If
        If LenB(viewText) = 0 Then
            viewText = segment
        Else
            viewText = viewText & vbLf & segment
        End If
    Next idx
    If startLine > 0 Then viewText = ChrW(8230) & viewText
    If endLine < UBound(hardLines) Then viewText = viewText & ChrW(8230)
    WindowFieldText = viewText
End Function

Private Function WindowLineAroundCaret( _
    ByVal lineText As String, _
    ByVal caretIndexInLine As Long, _
    ByVal charsPerLine As Long _
) As String
    Dim startChar As Long
    Dim windowText As String

    If Len(lineText) <= charsPerLine Then
        WindowLineAroundCaret = lineText
        Exit Function
    End If
    startChar = caretIndexInLine - charsPerLine + 2
    If startChar < 1 Then startChar = 1
    If startChar + charsPerLine - 1 > Len(lineText) Then
        startChar = Len(lineText) - charsPerLine + 1
    End If
    windowText = Mid$(lineText, startChar, charsPerLine)
    If startChar > 1 Then windowText = ChrW(8230) & windowText
    If startChar + charsPerLine - 1 < Len(lineText) Then
        windowText = windowText & ChrW(8230)
    End If
    WindowLineAroundCaret = windowText
End Function

Private Function ResolveSpinnerColor(ByVal themeValue As ReDimUI) As Long
    If mFillColor >= 0 Then
        ResolveSpinnerColor = mFillColor
    Else
        ResolveSpinnerColor = themeValue.PrimaryColor
    End If
End Function

Private Sub ApplyText(ByVal mainShape As Shape, ByVal themeValue As ReDimUI)
    Dim shownText As String
    Dim resolvedColor As Long
    Dim resolvedSize As Double
    Dim textChanged As Boolean

    If mKind = KIND_SPINNER Or mKind = KIND_TOGGLE Or mKind = KIND_OVERLAY _
        Or mKind = KIND_IMAGE Or IsCompositeDrawnKind() Then Exit Sub
    If (mKind = KIND_INPUT Or mKind = KIND_COMBO) And Not IsFloatField() Then
        Exit Sub
    End If
    shownText = ResolveText(Nothing)
    resolvedColor = ResolveTextColor(themeValue)
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    ' Shape text cannot scroll, so a focused field renders the tail
    ' window of its buffer: the caret lives at the end, and the view
    ' follows it. The full text stays in the buffer and commits intact.
    If IsFloatField() Then
        If ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
            shownText = WindowFieldText(shownText, resolvedSize)
        End If
    End If
    If Not mAppliedValid Then
        mainShape.TextFrame2.MarginLeft = 6
        mainShape.TextFrame2.MarginRight = 6
        mainShape.TextFrame2.MarginTop = 3
        mainShape.TextFrame2.MarginBottom = 3
        mainShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
        mainShape.TextFrame2.WordWrap = msoTrue
        If mKind = KIND_BUTTON Then
            mainShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Else
            mainShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End If
        If mKind = KIND_CARD Or mMultiLine Then
            mainShape.TextFrame2.VerticalAnchor = msoAnchorTop
        End If
    End If
    ' Measured, not assumed: replacing shape text preserves uniform
    ' character formatting, so the font block runs only when the style
    ' itself changed or the text crossed an empty transition (the one
    ' case the measurement did not cover). Steady-state typing is a
    ' single text write per keystroke.
    Dim mustRestyle As Boolean
    mustRestyle = (Not mAppliedValid) _
        Or mAppliedTextColor <> resolvedColor _
        Or mAppliedFontSize <> resolvedSize _
        Or mAppliedFontBold <> mFontBold
    textChanged = (Not mAppliedValid) Or (mAppliedText <> shownText)
    If textChanged Then
        If LenB(mAppliedText) = 0 Or LenB(shownText) = 0 Then
            mustRestyle = True
        End If
        mainShape.TextFrame2.TextRange.Text = shownText
        mAppliedText = shownText
    End If
    If mustRestyle Then
        With mainShape.TextFrame2.TextRange.Font
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
Private Sub ApplyBehavior(ByVal mainShape As Shape, ByVal hostSheet As Worksheet)
    Dim desiredAction As String

    desiredAction = "'" & hostSheet.Parent.Name & "'!RdxDispatch"
    If Not mAppliedValid Or mAppliedOnAction <> desiredAction Then
        mainShape.OnAction = desiredAction
        mAppliedOnAction = desiredAction
    End If
End Sub

' Composite and control-backed widgets keep their extra pieces in sync here:
' the toggle knob, the progress fill, form control values, and the input
' cell's one-time formatting. Part geometry is synced unconditionally; the
' shapes are tiny and correctness beats caching for moving pieces.
Private Sub ApplyKindExtras( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Select Case mKind
        Case KIND_TOGGLE
            ApplyToggleKnob hostSheet, themeValue
            mAppliedChecked = mChecked
        Case KIND_PROGRESS
            ' Tween loops repaint bars at frame rate; skip the fill
            ' part's shape writes when nothing about it would change.
            Dim progressSig As String
            progressSig = CStr(mValue) & "|" & CStr(ResolveFill(themeValue)) _
                & "|" & CStr(mVisible) & "|" & CStr(mLeft) & "|" & CStr(mTop) _
                & "|" & CStr(mWidth) & "|" & CStr(mHeight)
            If Not mAppliedValid Or mAppliedProgressSig <> progressSig Then
                ApplyProgressFill hostSheet, themeValue
                mAppliedValue = mValue
                mAppliedProgressSig = progressSig
            End If
        Case KIND_INPUT
            If Not IsFloatField() Then ApplyInputCell hostSheet, themeValue
        Case KIND_COMBO
            If Not IsFloatField() Then ApplyInputCell hostSheet, themeValue
            ApplyComboParts hostSheet, themeValue
        Case KIND_SELECT
            ApplySelectParts hostSheet, themeValue
        Case KIND_TICKBOX
            ApplyTickBox mainShape, hostSheet, themeValue
        Case KIND_RADIO
            ApplyRadioGroup mainShape, hostSheet, themeValue
        Case KIND_STEPPER
            ApplyStepper mainShape, hostSheet, themeValue
        Case KIND_SLIDEBAR
            ApplySlideBar mainShape, hostSheet, themeValue
        Case KIND_TRANSFER
            ApplyTransferParts mainShape, hostSheet, themeValue
        Case KIND_CHECKLIST
            ApplyCheckListParts mainShape, hostSheet, themeValue
        Case KIND_TOAST
            If mSlideRemaining > 0 Then mainShape.Top = mTop + mSlideRemaining
    End Select
End Sub

Private Sub StyleGlyphText( _
    ByVal glyphShape As Shape, _
    ByVal glyph As String, _
    ByVal inkColor As Long, _
    ByVal glyphSize As Double, _
    ByVal themeValue As ReDimUI _
)
    With glyphShape.TextFrame2
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
    ByVal captionShape As Shape, _
    ByVal captionText As String, _
    ByVal themeValue As ReDimUI _
)
    Dim resolvedSize As Double

    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    captionShape.Fill.Visible = msoFalse
    captionShape.Line.Visible = msoFalse
    With captionShape.TextFrame2
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
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim captionShape As Shape
    Dim boxSize As Double

    boxSize = mHeight - 2
    If boxSize > 16 Then boxSize = 16
    If boxSize < 10 Then boxSize = 10
    mainShape.Left = mLeft
    mainShape.Top = mTop + (mHeight - boxSize) / 2
    mainShape.Width = boxSize
    mainShape.Height = boxSize
    mainShape.Visible = BoolToMso(mVisible)
    mainShape.Fill.Visible = msoTrue
    If mChecked And mEnabled Then
        mainShape.Fill.ForeColor.RGB = ResolveFill(themeValue)
        mainShape.Line.Visible = msoFalse
    Else
        mainShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        mainShape.Line.Visible = msoTrue
        mainShape.Line.ForeColor.RGB = themeValue.BorderColor
        mainShape.Line.Weight = 1.25
    End If
    If mChecked Then
        StyleGlyphText mainShape, ChrW(10003), _
            IIf(mEnabled, themeValue.OnPrimaryColor, themeValue.OnMutedColor), _
            boxSize * 0.62, themeValue
    Else
        StyleGlyphText mainShape, vbNullString, themeValue.OnSurfaceColor, _
            boxSize * 0.62, themeValue
    End If
    Set captionShape = EnsurePartShape(hostSheet, "lbl", msoShapeRectangle)
    captionShape.Left = mLeft + boxSize + 6
    captionShape.Top = mTop - 1
    captionShape.Width = mWidth - boxSize - 6
    captionShape.Height = mHeight + 2
    StyleCaptionPart captionShape, mText, themeValue
    captionShape.Visible = BoolToMso(mVisible)
    captionShape.OnAction = mAppliedOnAction
    mAppliedChecked = mChecked
End Sub

' One circle-and-caption row per item. The selected row's circle fills with
' the accent and carries an inner dot part.
Private Sub ApplyRadioGroup( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim rowTotal As Long
    Dim rowPitch As Double
    Dim rowTop As Double
    Dim idx As Long
    Dim ring As Shape
    Dim captionShape As Shape
    Dim dot As Shape
    Dim circleSize As Double
    Dim isPicked As Boolean

    If mItems Is Nothing Then Exit Sub
    rowTotal = mItems.Count
    If rowTotal = 0 Then Exit Sub
    rowPitch = mHeight / rowTotal
    circleSize = rowPitch - 6
    If circleSize > 14 Then circleSize = 14
    If circleSize < 10 Then circleSize = 10
    For idx = 1 To rowTotal
        rowTop = mTop + (idx - 1) * rowPitch
        isPicked = (idx = CLng(mValue))
        If idx = 1 Then
            Set ring = mainShape
        Else
            Set ring = EnsurePartShape(hostSheet, "c" & idx, msoShapeOval)
        End If
        ring.Left = mLeft
        ring.Top = rowTop + (rowPitch - circleSize) / 2
        ring.Width = circleSize
        ring.Height = circleSize
        ring.Fill.Visible = msoTrue
        ring.Fill.ForeColor.RGB = themeValue.SurfaceColor
        ring.Line.Visible = msoTrue
        If isPicked And mEnabled Then
            ring.Line.ForeColor.RGB = ResolveFill(themeValue)
            ring.Line.Weight = 1.5
        Else
            ring.Line.ForeColor.RGB = themeValue.BorderColor
            ring.Line.Weight = 1.25
        End If
        ring.Visible = BoolToMso(mVisible)
        ring.OnAction = mAppliedOnAction
        Set dot = EnsurePartShape(hostSheet, "d" & idx, msoShapeOval)
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
        dot.Visible = BoolToMso(mVisible And isPicked)
        dot.OnAction = mAppliedOnAction
        Set captionShape = EnsurePartShape(hostSheet, "t" & idx, msoShapeRectangle)
        captionShape.Left = mLeft + circleSize + 8
        captionShape.Top = rowTop
        captionShape.Width = mWidth - circleSize - 8
        captionShape.Height = rowPitch
        StyleCaptionPart captionShape, ItemAt(idx), themeValue
        captionShape.Visible = BoolToMso(mVisible)
        captionShape.OnAction = mAppliedOnAction
    Next idx
    ' Rows beyond the current item count are leftovers from a longer list.
    For idx = rowTotal + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "c" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "d" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & idx).Delete
        On Error GoTo 0
    Next idx
    mAppliedOptionCount = rowTotal
    mAppliedValue = mValue
End Sub

' One box-and-caption row per item, any number checked, plus an optional
' select-all header whose box reads empty, checked, or dashed for mixed.
Private Sub ApplyCheckListParts( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim rowTotal As Long
    Dim headerRows As Long
    Dim rowPitch As Double
    Dim rowTop As Double
    Dim idx As Long
    Dim box As Shape
    Dim captionShape As Shape
    Dim boxSize As Double
    Dim isOn As Boolean
    Dim checkedTotal As Long

    EnsureCheckedSet
    If mItems Is Nothing Then Exit Sub
    rowTotal = mItems.Count
    If rowTotal = 0 Then Exit Sub
    PruneTransferSelection mCheckedSet, rowTotal
    ' Skip the shape work when nothing visible changed.
    Dim checkSig As String
    checkSig = CStr(ObjPtr(themeValue)) & ":" & CStr(mLeft) & ":" & _
        CStr(mTop) & ":" & CStr(mWidth) & ":" & CStr(mHeight) & ":" & _
        CStr(mVisible) & ":" & CStr(mEnabled) & ":" & _
        CStr(mSelectAllRow) & ":" & CStr(mFontSize) & ":" & _
        JoinedSig(mItems) & "#" & JoinedSig(mCheckedSet)
    If mAppliedValid And checkSig = mAppliedCheckSig Then Exit Sub
    mAppliedCheckSig = checkSig
    checkedTotal = mCheckedSet.Count
    headerRows = 0
    If mSelectAllRow Then headerRows = 1
    rowPitch = mHeight / (rowTotal + headerRows)
    boxSize = rowPitch - 6
    If boxSize > 14 Then boxSize = 14
    If boxSize < 10 Then boxSize = 10

    If mSelectAllRow Then
        Set box = EnsurePartShape(hostSheet, "mb", msoShapeRoundedRectangle)
        On Error Resume Next
        box.Adjustments(1) = 0.2
        On Error GoTo 0
        StyleCheckBoxFace box, boxSize, mTop, rowPitch, themeValue, _
            (checkedTotal = rowTotal), (checkedTotal > 0)
        Set captionShape = EnsurePartShape(hostSheet, "mt", msoShapeRectangle)
        captionShape.Left = mLeft + boxSize + 8
        captionShape.Top = mTop
        captionShape.Width = mWidth - boxSize - 8
        captionShape.Height = rowPitch
        StyleCaptionPart captionShape, _
            "Select all (" & checkedTotal & "/" & rowTotal & ")", themeValue
        captionShape.Visible = BoolToMso(mVisible)
        captionShape.OnAction = mAppliedOnAction
    Else
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "mb").Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "mt").Delete
        On Error GoTo 0
    End If

    For idx = 1 To rowTotal
        rowTop = mTop + (idx - 1 + headerRows) * rowPitch
        isOn = TransferSelectionHas(mCheckedSet, idx)
        If idx = 1 Then
            Set box = mainShape
        Else
            Set box = EnsurePartShape( _
                hostSheet, "b" & idx, msoShapeRoundedRectangle)
            On Error Resume Next
            box.Adjustments(1) = 0.2
            On Error GoTo 0
        End If
        StyleCheckBoxFace box, boxSize, rowTop, rowPitch, themeValue, _
            isOn, isOn
        Set captionShape = EnsurePartShape( _
            hostSheet, "t" & idx, msoShapeRectangle)
        captionShape.Left = mLeft + boxSize + 8
        captionShape.Top = rowTop
        captionShape.Width = mWidth - boxSize - 8
        captionShape.Height = rowPitch
        StyleCaptionPart captionShape, ItemAt(idx), themeValue
        captionShape.Visible = BoolToMso(mVisible)
        captionShape.OnAction = mAppliedOnAction
    Next idx
    For idx = rowTotal + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "b" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & idx).Delete
        On Error GoTo 0
    Next idx
    mAppliedOptionCount = rowTotal
End Sub

' A checkbox face in one of three reads: full (accent fill and a check),
' partial (accent fill and a dash, the mixed state), or empty (surface
' fill and a border). full implies partial.
Private Sub StyleCheckBoxFace( _
    ByVal box As Shape, _
    ByVal boxSize As Double, _
    ByVal rowTop As Double, _
    ByVal rowPitch As Double, _
    ByVal themeValue As ReDimUI, _
    ByVal fullState As Boolean, _
    ByVal partialState As Boolean _
)
    box.Left = mLeft
    box.Top = rowTop + (rowPitch - boxSize) / 2
    box.Width = boxSize
    box.Height = boxSize
    box.Fill.Visible = msoTrue
    If partialState And mEnabled Then
        box.Fill.ForeColor.RGB = ResolveFill(themeValue)
        box.Line.Visible = msoFalse
    Else
        box.Fill.ForeColor.RGB = themeValue.SurfaceColor
        box.Line.Visible = msoTrue
        box.Line.ForeColor.RGB = themeValue.BorderColor
        box.Line.Weight = 1.25
    End If
    If fullState Then
        StyleGlyphText box, ChrW(10003), _
            IIf(mEnabled, themeValue.OnPrimaryColor, themeValue.OnMutedColor), _
            boxSize * 0.62, themeValue
    ElseIf partialState Then
        StyleGlyphText box, "-", _
            IIf(mEnabled, themeValue.OnPrimaryColor, themeValue.OnMutedColor), _
            boxSize * 0.8, themeValue
    Else
        StyleGlyphText box, vbNullString, themeValue.OnSurfaceColor, _
            boxSize * 0.62, themeValue
    End If
    box.Visible = BoolToMso(mVisible)
    box.OnAction = mAppliedOnAction
End Sub

' Minus and plus squares around a bordered value face.
Private Sub ApplyStepper( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim minusButton As Shape
    Dim plusButton As Shape
    Dim buttonSize As Double

    buttonSize = mHeight
    mainShape.Left = mLeft + buttonSize + 4
    mainShape.Top = mTop
    mainShape.Width = mWidth - 2 * (buttonSize + 4)
    mainShape.Height = mHeight
    mainShape.Visible = BoolToMso(mVisible)
    mainShape.Fill.Visible = msoTrue
    mainShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
    mainShape.Line.Visible = msoTrue
    mainShape.Line.ForeColor.RGB = themeValue.BorderColor
    mainShape.Line.Weight = 1
    StyleGlyphText mainShape, CStr(CLng(mValue)), ResolveDrawnInk(themeValue), _
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

' Dual listbox: the main shape is the left panel; the right panel, two
' counted headers, four move buttons, and the row shapes are parts. Rows
' render up to the panel capacity; the header count keeps overflow honest.
Private Sub ApplyTransferParts( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Const HEADER_H As Double = 18
    Const BTN_W As Double = 26
    Const BTN_H As Double = 22
    Const PANEL_GAP As Double = 6
    Dim panelW As Double
    Dim panelH As Double
    Dim panelTop As Double
    Dim rightX As Double
    Dim btnX As Double
    Dim btnTop As Double
    Dim rightPanel As Shape
    Dim headerShape As Shape

    EnsureTransferCollections
    PruneTransferSelection mSelLeft, mItems.Count
    PruneTransferSelection mSelRight, mChosenItems.Count
    If mLeftOffset > mItems.Count - TransferRowCapacity() Then
        mLeftOffset = mItems.Count - TransferRowCapacity()
    End If
    If mLeftOffset < 0 Then mLeftOffset = 0
    If mRightOffset > mChosenItems.Count - TransferRowCapacity() Then
        mRightOffset = mChosenItems.Count - TransferRowCapacity()
    End If
    If mRightOffset < 0 Then mRightOffset = 0
    ' A repaint that changes nothing visible skips all the shape work:
    ' both panels, headers, rows, and buttons carry dozens of COM
    ' writes, and most touches change only one side or nothing.
    Dim transferSig As String
    transferSig = CStr(ObjPtr(themeValue)) & ":" & CStr(mLeft) & ":" & _
        CStr(mTop) & ":" & CStr(mWidth) & ":" & CStr(mHeight) & ":" & _
        CStr(mVisible) & ":" & mCaptionLeft & "/" & mCaptionRight & ":" & _
        CStr(mLeftOffset) & ":" & CStr(mRightOffset) & ":" & _
        CStr(mFontSize) & ":" & _
        JoinedSig(mItems) & "#" & JoinedSig(mChosenItems) & "#" & _
        JoinedSig(mSelLeft) & "#" & JoinedSig(mSelRight)
    If mAppliedValid And transferSig = mAppliedTransferSig Then Exit Sub
    mAppliedTransferSig = transferSig
    panelW = (mWidth - BTN_W - 4 * PANEL_GAP) / 2
    If panelW < 40 Then panelW = 40
    panelTop = mTop + HEADER_H
    panelH = mHeight - HEADER_H
    If panelH < 30 Then panelH = 30
    rightX = mLeft + panelW + BTN_W + 4 * PANEL_GAP
    btnX = mLeft + panelW + 2 * PANEL_GAP

    mainShape.Left = mLeft
    mainShape.Top = panelTop
    mainShape.Width = panelW
    mainShape.Height = panelH
    mainShape.Fill.Visible = msoTrue
    mainShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
    mainShape.Line.Visible = msoTrue
    mainShape.Line.ForeColor.RGB = themeValue.BorderColor
    mainShape.Line.Weight = 1
    mainShape.Visible = BoolToMso(mVisible)

    Set rightPanel = EnsurePartShape(hostSheet, "rp", msoShapeRoundedRectangle)
    On Error Resume Next
    rightPanel.Adjustments(1) = 0.08
    On Error GoTo 0
    rightPanel.Left = rightX
    rightPanel.Top = panelTop
    rightPanel.Width = panelW
    rightPanel.Height = panelH
    rightPanel.Fill.Visible = msoTrue
    rightPanel.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
    rightPanel.Line.Visible = msoTrue
    rightPanel.Line.ForeColor.RGB = themeValue.BorderColor
    rightPanel.Line.Weight = 1
    rightPanel.Visible = BoolToMso(mVisible)
    rightPanel.OnAction = mAppliedOnAction

    Set headerShape = EnsurePartShape(hostSheet, "hl", msoShapeRectangle)
    headerShape.Left = mLeft
    headerShape.Top = mTop - 1
    headerShape.Width = panelW
    headerShape.Height = HEADER_H
    StyleCaptionPart headerShape, _
        IIf(LenB(mCaptionLeft) > 0, mCaptionLeft, "Available") & _
        " (" & mItems.Count & ")", themeValue
    headerShape.Visible = BoolToMso(mVisible)

    Set headerShape = EnsurePartShape(hostSheet, "hr", msoShapeRectangle)
    headerShape.Left = rightX
    headerShape.Top = mTop - 1
    headerShape.Width = panelW
    headerShape.Height = HEADER_H
    StyleCaptionPart headerShape, _
        IIf(LenB(mCaptionRight) > 0, mCaptionRight, "Selected") & _
        " (" & mChosenItems.Count & ")", themeValue
    headerShape.Visible = BoolToMso(mVisible)

    btnTop = panelTop + (panelH - (4 * BTN_H + 3 * PANEL_GAP)) / 2
    If btnTop < panelTop Then btnTop = panelTop
    ApplyTransferButton hostSheet, themeValue, "mvr", ">", _
        btnX, btnTop, BTN_W, BTN_H
    ApplyTransferButton hostSheet, themeValue, "mvar", ">>", _
        btnX, btnTop + (BTN_H + PANEL_GAP), BTN_W, BTN_H
    ApplyTransferButton hostSheet, themeValue, "mvl", "<", _
        btnX, btnTop + 2 * (BTN_H + PANEL_GAP), BTN_W, BTN_H
    ApplyTransferButton hostSheet, themeValue, "mval", "<<", _
        btnX, btnTop + 3 * (BTN_H + PANEL_GAP), BTN_W, BTN_H

    SyncTransferRows hostSheet, themeValue, mItems, "al", _
        mLeft, panelTop, panelW, panelH, mSelLeft, mLeftOffset, _
        mAppliedLeftRows
    SyncTransferRows hostSheet, themeValue, mChosenItems, "cl", _
        rightX, panelTop, panelW, panelH, mSelRight, mRightOffset, _
        mAppliedRightRows
    ApplyTransferScroller hostSheet, themeValue, "alu", "ald", _
        mLeft, panelTop, panelW, panelH, mItems.Count
    ApplyTransferScroller hostSheet, themeValue, "clu", "cld", _
        rightX, panelTop, panelW, panelH, mChosenItems.Count
End Sub

' Small paging arrows on a panel's right edge, shown only when the list
' outgrows the panel.
Private Sub ApplyTransferScroller( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal upPart As String, _
    ByVal downPart As String, _
    ByVal panelX As Double, _
    ByVal panelTop As Double, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal itemTotal As Long _
)
    Dim scrollButton As Shape

    If itemTotal > TransferRowCapacity() And mVisible Then
        Set scrollButton = EnsurePartShape( _
            hostSheet, upPart, msoShapeRoundedRectangle)
        scrollButton.Left = panelX + panelW - 17
        scrollButton.Top = panelTop + 3
        scrollButton.Width = 14
        scrollButton.Height = 14
        scrollButton.Fill.Visible = msoTrue
        scrollButton.Fill.ForeColor.RGB = themeValue.MutedColor
        StyleGlyphText scrollButton, ChrW(9650), _
            ResolveDrawnInk(themeValue), 7, themeValue
        scrollButton.Visible = msoTrue
        scrollButton.OnAction = mAppliedOnAction
        scrollButton.ZOrder msoBringToFront
        Set scrollButton = EnsurePartShape( _
            hostSheet, downPart, msoShapeRoundedRectangle)
        scrollButton.Left = panelX + panelW - 17
        scrollButton.Top = panelTop + panelH - 17
        scrollButton.Width = 14
        scrollButton.Height = 14
        scrollButton.Fill.Visible = msoTrue
        scrollButton.Fill.ForeColor.RGB = themeValue.MutedColor
        StyleGlyphText scrollButton, ChrW(9660), _
            ResolveDrawnInk(themeValue), 7, themeValue
        scrollButton.Visible = msoTrue
        scrollButton.OnAction = mAppliedOnAction
        scrollButton.ZOrder msoBringToFront
    Else
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & upPart).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & downPart).Delete
        On Error GoTo 0
    End If
End Sub

Private Sub ApplyTransferButton( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal glyph As String, _
    ByVal buttonLeft As Double, _
    ByVal buttonTop As Double, _
    ByVal buttonWidth As Double, _
    ByVal buttonHeight As Double _
)
    Dim moveButton As Shape

    Set moveButton = EnsurePartShape( _
        hostSheet, partName, msoShapeRoundedRectangle)
    On Error Resume Next
    moveButton.Adjustments(1) = 0.25
    On Error GoTo 0
    moveButton.Left = buttonLeft
    moveButton.Top = buttonTop
    moveButton.Width = buttonWidth
    moveButton.Height = buttonHeight
    moveButton.Fill.Visible = msoTrue
    moveButton.Fill.ForeColor.RGB = themeValue.MutedColor
    StyleGlyphText moveButton, glyph, ResolveDrawnInk(themeValue), _
        themeValue.BaseFontSize, themeValue
    moveButton.Visible = BoolToMso(mVisible)
    moveButton.OnAction = mAppliedOnAction
End Sub

Private Sub SyncTransferRows( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal itemsRef As Collection, _
    ByVal partPrefix As String, _
    ByVal panelX As Double, _
    ByVal panelTop As Double, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal selectionSet As Collection, _
    ByVal scrollOffset As Long, _
    ByRef appliedRows As Long _
)
    Const ROW_H As Double = 20
    Dim rowShape As Shape
    Dim idx As Long
    Dim wanted As Long
    Dim itemIndex As Long
    Dim rowWidth As Double
    Dim resolvedSize As Double
    Dim rowSelected As Boolean

    wanted = itemsRef.Count - scrollOffset
    If wanted > Int((panelH - 6) / ROW_H) Then
        wanted = Int((panelH - 6) / ROW_H)
    End If
    If wanted < 0 Then wanted = 0
    If Not mVisible Then wanted = 0
    rowWidth = panelW - 6
    If itemsRef.Count > Int((panelH - 6) / ROW_H) Then
        ' Leave room for the paging arrows.
        rowWidth = rowWidth - 16
    End If
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    For idx = 1 To wanted
        itemIndex = scrollOffset + idx
        rowSelected = TransferSelectionHas(selectionSet, itemIndex)
        Set rowShape = EnsurePartShape( _
            hostSheet, partPrefix & idx, msoShapeRectangle)
        rowShape.Left = panelX + 3
        rowShape.Top = panelTop + 3 + (idx - 1) * ROW_H
        rowShape.Width = rowWidth
        rowShape.Height = ROW_H
        rowShape.Fill.Visible = msoTrue
        If rowSelected Then
            rowShape.Fill.ForeColor.RGB = themeValue.PrimaryColor
        Else
            rowShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
        End If
        With rowShape.TextFrame2
            .MarginLeft = 6
            .MarginRight = 4
            .MarginTop = 1
            .MarginBottom = 1
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .TextRange.Text = CStr(itemsRef.Item(itemIndex))
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            With .TextRange.Font
                .Name = themeValue.FontName
                .Size = resolvedSize
                .Bold = msoFalse
                If rowSelected Then
                    .Fill.ForeColor.RGB = themeValue.OnPrimaryColor
                Else
                    .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
                End If
            End With
        End With
        rowShape.Visible = msoTrue
        rowShape.OnAction = mAppliedOnAction
        rowShape.ZOrder msoBringToFront
    Next idx
    For idx = wanted + 1 To appliedRows
        On Error Resume Next
        hostSheet.Shapes( _
            ShapeName & PART_SEPARATOR & partPrefix & idx).Delete
        On Error GoTo 0
    Next idx
    appliedRows = wanted
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
    SyncDropList hostSheet, themeValue
End Sub

' Slim centered track, accent fill up to the value, and a round thumb.
Private Sub ApplySlideBar( _
    ByVal mainShape As Shape, _
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
    mainShape.Left = mLeft
    mainShape.Top = trackTop
    mainShape.Width = mWidth
    mainShape.Height = 6
    mainShape.Visible = BoolToMso(mVisible)
    mainShape.Fill.Visible = msoTrue
    mainShape.Fill.ForeColor.RGB = themeValue.MutedColor
    mainShape.Line.Visible = msoFalse

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
            ' Floating chrome paints above the track: a press that lands
            ' on an open drop list or a modal overlay belongs to the
            ' chrome, and the geometric track test must not reach
            ' through it.
            If Not PressClaimedByChrome(hostApp) Then
                If CursorOverTrack() Then
                    fraction = CursorTrackFraction()
                    If fraction >= 0 Then BeginSlideDrag hostApp, fraction
                End If
            End If
        End If
    End If
    mSlidePrevButtonDown = buttonDown
End Sub

Private Function PressClaimedByChrome(ByVal hostApp As ReDimUI) As Boolean
    Dim screenPoint As RdxScreenPoint
    Dim pointsX As Double
    Dim pointsY As Double

    If GetCursorPos(screenPoint) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(screenPoint.pixelX)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(screenPoint.pixelY)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    PressClaimedByChrome = hostApp.PointClaimedByChrome( _
        pointsX, pointsY, mComponentId)
End Function

' True when this component projects floating chrome over the given
' sheet point: a visible modal overlay claims everything it covers, an
' open drop list claims its face plus the dropped rows and pagers.
Friend Function ChromeClaimsPoint( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double _
) As Boolean
    Dim listDepth As Double

    If Not mVisible Then Exit Function
    If mKind = KIND_OVERLAY Then
        If pointsX >= mLeft And pointsX <= mLeft + mWidth _
            And pointsY >= mTop And pointsY <= mTop + mHeight Then
            ChromeClaimsPoint = True
        End If
        Exit Function
    End If
    If (mKind = KIND_COMBO Or mKind = KIND_SELECT) And mSelectOpen Then
        listDepth = DropListDepth()
        If pointsX >= mLeft - 2 And pointsX <= mLeft + mWidth + 2 _
            And pointsY >= mTop _
            And pointsY <= mTop + mHeight + listDepth + 2 Then
            ChromeClaimsPoint = True
        End If
    End If
End Function

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
    Dim screenPoint As RdxScreenPoint
    Dim pointsX As Double
    Dim pointsY As Double

    If GetCursorPos(screenPoint) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(screenPoint.pixelX)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(screenPoint.pixelY)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    If pointsX < mLeft - 8 Or pointsX > mLeft + mWidth + 8 Then Exit Function
    If pointsY < mTop - 4 Or pointsY > mTop + mHeight + 4 Then Exit Function
    CursorOverTrack = True
End Function

Private Function CursorTrackFraction() As Double
    Dim screenPoint As RdxScreenPoint
    Dim cursorPoints As Double

    CursorTrackFraction = -1
    If GetCursorPos(screenPoint) = 0 Then Exit Function
    cursorPoints = ScreenXToSheetPoints(screenPoint.pixelX)
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
    Dim pixelsPerPoint As Double

    ScreenXToSheetPoints = UNMAPPABLE_POINTS
    On Error GoTo Unmappable
    originPx = ActiveWindow.PointsToScreenPixelsX(0)
    referencePx = ActiveWindow.PointsToScreenPixelsX(1000)
    measuredSlope = (referencePx - originPx) / 1000#
    dpiSlope = ScreenPixelsPerPoint()
    On Error GoTo 0
    If measuredSlope <= 0 Or dpiSlope <= 0 Then Exit Function
    If Abs(measuredSlope - dpiSlope) < 0.02 Then
        pixelsPerPoint = measuredSlope
    ElseIf Abs(measuredSlope - 1#) < 0.02 Then
        pixelsPerPoint = dpiSlope
    Else
        pixelsPerPoint = measuredSlope
    End If
    ScreenXToSheetPoints = (screenX - originPx) / pixelsPerPoint
    Exit Function

Unmappable:
    ScreenXToSheetPoints = UNMAPPABLE_POINTS
End Function

Friend Function ScreenYToSheetPoints(ByVal screenY As Long) As Double
    Dim originPx As Double
    Dim referencePx As Double
    Dim measuredSlope As Double
    Dim dpiSlope As Double
    Dim pixelsPerPoint As Double

    ScreenYToSheetPoints = UNMAPPABLE_POINTS
    On Error GoTo Unmappable
    originPx = ActiveWindow.PointsToScreenPixelsY(0)
    referencePx = ActiveWindow.PointsToScreenPixelsY(1000)
    measuredSlope = (referencePx - originPx) / 1000#
    dpiSlope = ScreenPixelsPerPointY()
    On Error GoTo 0
    If measuredSlope <= 0 Or dpiSlope <= 0 Then Exit Function
    If Abs(measuredSlope - dpiSlope) < 0.02 Then
        pixelsPerPoint = measuredSlope
    ElseIf Abs(measuredSlope - 1#) < 0.02 Then
        pixelsPerPoint = dpiSlope
    Else
        pixelsPerPoint = measuredSlope
    End If
    ScreenYToSheetPoints = (screenY - originPx) / pixelsPerPoint
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
    SyncDropList hostSheet, themeValue
End Sub

' The open drop list of a SelectBox or ComboBox: at most mListRows item
' rows from mListOffset, with a clickable pager row at each edge that has
' items beyond it. Rows map back to item indexes through mListMatches.
' The combo's keyboard highlight paints in the accent, the current item
' in the muted fill.
Private Sub SyncDropList( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim optionShape As Shape
    Dim idx As Long
    Dim wanted As Long
    Dim itemIndex As Long
    Dim resolvedSize As Double
    Dim matchTotal As Long
    Dim aboveCount As Long
    Dim belowCount As Long
    Dim rowShift As Long
    Dim matchIndex As Long
    Dim rowPitch As Double

    Set mListMatches = BuildListMatches(hostSheet)
    matchTotal = mListMatches.Count
    If mListOffset > matchTotal - mListRows Then
        mListOffset = matchTotal - mListRows
    End If
    If mListOffset < 0 Then mListOffset = 0
    If mListHighlight > matchTotal Then mListHighlight = 0
    If mSelectOpen And mVisible Then
        wanted = matchTotal - mListOffset
        If wanted > mListRows Then wanted = mListRows
        aboveCount = mListOffset
        belowCount = matchTotal - mListOffset - wanted
    End If
    If aboveCount > 0 Then rowShift = 1
    rowPitch = DropRowPitch()
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    ' Keystrokes and repaints that leave the visible list identical (same
    ' window, rows, highlight, selection, geometry, and theme) skip the
    ' shape work entirely.
    Dim listSig As String
    listSig = CStr(ObjPtr(themeValue)) & ":" & CStr(wanted) & ":" & _
        CStr(mValue) & ":" & CStr(mListOffset) & ":" & _
        CStr(mListHighlight) & ":" & CStr(aboveCount) & ":" & _
        CStr(belowCount) & ":" & CStr(mLeft) & ":" & CStr(mTop) & ":" & _
        CStr(mWidth) & ":" & CStr(mHeight) & ":" & CStr(resolvedSize)
    For idx = 1 To wanted
        itemIndex = CLng(mListMatches.Item(mListOffset + idx))
        listSig = listSig & "|" & CStr(itemIndex) & "=" & ItemAt(itemIndex)
    Next idx
    If mAppliedValid And listSig = mAppliedListSig Then Exit Sub
    mAppliedListSig = listSig
    SyncListPager hostSheet, themeValue, "optu", ChrW(9650), _
        aboveCount, mTop + mHeight, resolvedSize
    For idx = 1 To wanted
        matchIndex = mListOffset + idx
        itemIndex = CLng(mListMatches.Item(matchIndex))
        Set optionShape = EnsurePartShape( _
            hostSheet, "opt" & idx, msoShapeRectangle)
        optionShape.Left = mLeft
        optionShape.Top = mTop + mHeight + (rowShift + idx - 1) * rowPitch
        optionShape.Width = mWidth
        optionShape.Height = rowPitch
        optionShape.Fill.Visible = msoTrue
        If matchIndex = mListHighlight Then
            optionShape.Fill.ForeColor.RGB = themeValue.PrimaryColor
        ElseIf itemIndex = CLng(mValue) Then
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
                If matchIndex = mListHighlight Then
                    .Fill.ForeColor.RGB = themeValue.OnPrimaryColor
                Else
                    .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
                End If
            End With
        End With
        optionShape.Visible = msoTrue
        optionShape.OnAction = mAppliedOnAction
        optionShape.ZOrder msoBringToFront
    Next idx
    For idx = wanted + 1 To mAppliedOptionCount
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & idx).Delete
        On Error GoTo 0
    Next idx
    mAppliedOptionCount = wanted
    SyncListPager hostSheet, themeValue, "optd", ChrW(9660), _
        belowCount, mTop + mHeight + (rowShift + wanted) * rowPitch, _
        resolvedSize
    ' The 0.17.0 inert indicator row is superseded by the pagers.
    On Error Resume Next
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "optm").Delete
    On Error GoTo 0
End Sub

' A SelectBox row is exactly the face height; a combo row runs two
' points taller.
Private Function DropRowPitch() As Double
    If mKind = KIND_COMBO Then
        DropRowPitch = mHeight + 2
    Else
        DropRowPitch = mHeight
    End If
End Function

' How far the open drop list reaches below the face: the drawn window of
' rows plus a pager row at each edge with items beyond it. Zero while the
' list is closed or not yet drawn.
Private Function DropListDepth() As Double
    Dim matchTotal As Long
    Dim shownRows As Long
    Dim rowShift As Long

    If Not mSelectOpen Or mListMatches Is Nothing Then Exit Function
    matchTotal = mListMatches.Count
    shownRows = matchTotal - mListOffset
    If shownRows > mListRows Then shownRows = mListRows
    If shownRows < 0 Then shownRows = 0
    If mListOffset > 0 Then rowShift = 1
    DropListDepth = (rowShift + shownRows) * DropRowPitch()
    If matchTotal - mListOffset - shownRows > 0 Then
        DropListDepth = DropListDepth + mHeight
    End If
End Function

' The item indexes the open list walks, in order. A SelectBox lists every
' item. A combo filters by its text (contains-match, case-insensitive),
' except while its list is pristine and the text names an item exactly:
' a picked or committed value reopens onto the whole list instead of
' trapping the user in its one match.
Private Function BuildListMatches(ByVal hostSheet As Worksheet) As Collection
    Dim filterText As String

    If mKind = KIND_COMBO Then
        filterText = FieldText(hostSheet)
        If mFilterPristine Then
            If ExactItemMatch(filterText) > 0 Then filterText = vbNullString
        End If
    End If
    Set BuildListMatches = FilterItems(filterText)
End Function

' Opens the list pristine, scrolled so the current item (the SelectBox
' selection, or the item a combo's text names) is the first row; the
' render clamps the window near the end of the list.
Private Sub OpenDropList(ByVal hostSheet As Worksheet)
    Dim currentItem As Long

    mSelectOpen = True
    mFilterPristine = True
    mListHighlight = 0
    If mKind = KIND_COMBO Then
        currentItem = ExactItemMatch(FieldText(hostSheet))
    Else
        currentItem = CLng(mValue)
    End If
    ' A pristine list holds every item, so an item's row is its index.
    If currentItem >= 1 And currentItem <= ItemCount Then
        mListOffset = currentItem - 1
    Else
        mListOffset = 0
    End If
End Sub

Private Sub ToggleDropList(ByVal hostSheet As Worksheet)
    If mSelectOpen Then
        mSelectOpen = False
    Else
        OpenDropList hostSheet
    End If
    Touch
End Sub

' A pager row pages the window by its height less one row of overlap and
' keeps the list open. The combo's keyboard highlight resets so the next
' arrow press starts inside the visible window.
Private Sub PageDropList(ByVal pageDown As Boolean)
    Dim pageRows As Long

    pageRows = mListRows - 1
    If pageRows < 1 Then pageRows = 1
    If pageDown Then
        mListOffset = mListOffset + pageRows
    Else
        mListOffset = mListOffset - pageRows
        If mListOffset < 0 Then mListOffset = 0
    End If
    mListHighlight = 0
    Touch
End Sub

' The item index a clicked window row stands for, or 0 for a row the
' list no longer draws.
Private Function ListItemAtRow(ByVal windowRow As Long) As Long
    Dim matchIndex As Long

    If mListMatches Is Nothing Then Exit Function
    matchIndex = mListOffset + windowRow
    If matchIndex < 1 Or matchIndex > mListMatches.Count Then Exit Function
    ListItemAtRow = CLng(mListMatches.Item(matchIndex))
End Function

' A clickable paging row at a list edge: the arrow plus the count of
' items beyond that edge. Clicking pages the window; absent when nothing
' lies beyond.
Private Sub SyncListPager( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal arrowGlyph As String, _
    ByVal beyondCount As Long, _
    ByVal rowTop As Double, _
    ByVal resolvedSize As Double _
)
    Dim pagerShape As Shape

    If beyondCount > 0 Then
        Set pagerShape = EnsurePartShape( _
            hostSheet, partName, msoShapeRectangle)
        pagerShape.Left = mLeft
        pagerShape.Top = rowTop
        pagerShape.Width = mWidth
        pagerShape.Height = mHeight
        pagerShape.Fill.Visible = msoTrue
        pagerShape.Fill.ForeColor.RGB = themeValue.MutedColor
        pagerShape.Line.Visible = msoTrue
        pagerShape.Line.ForeColor.RGB = themeValue.BorderColor
        pagerShape.Line.Weight = 1
        StyleGlyphText pagerShape, arrowGlyph & " " & _
            CStr(beyondCount) & " more", themeValue.OnSurfaceColor, _
            resolvedSize - 1, themeValue
        pagerShape.Visible = msoTrue
        pagerShape.OnAction = mAppliedOnAction
        pagerShape.ZOrder msoBringToFront
    Else
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & partName).Delete
        On Error GoTo 0
    End If
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
    mFieldSelAtFocus = CurrentSelectionKey()
    mCaretPos = Len(mText)
    mCaretOn = True
    mCaretAccumMs = 0
    ReDimUI.SetKeyboardFocus hostApp, Me
    If mKind = KIND_COMBO Then OpenDropList hostApp.Sheet
    Touch
End Sub

' The grid selection as a comparable key. A non-range selection (a shape)
' reads as the sentinel, so only real cell selections participate.
Private Function CurrentSelectionKey() As String
    On Error Resume Next
    CurrentSelectionKey = "?"
    CurrentSelectionKey = ActiveWindow.ActiveSheet.Name & "!" & _
        Selection.Address
    On Error GoTo 0
End Function

' Commit-blur keeps the text and fires events; revert-blur restores what
' was there when focus arrived and fires nothing.
Friend Sub BlurField(ByVal hostApp As ReDimUI, ByVal keepEdits As Boolean)
    Dim exactAt As Long

    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    ReDimUI.ClearKeyboardFocus
    mCaretOn = False
    If keepEdits Then
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
    Dim textBeforeKey As String

    If Not IsFloatField() Then Exit Sub
    If mCaretPos > Len(mText) Then mCaretPos = Len(mText)
    textBeforeKey = mText
    Select Case keyText
        Case "{ENTER}"
            ' Single-line: Enter commits - except a combo with a
            ' highlighted option, which takes that option. Multi-line:
            ' Enter is a newline and committing moves to Tab,
            ' Ctrl+Enter, or clicking away.
            If mMultiLine Then
                InsertAtCaret vbLf
            ElseIf mKind = KIND_COMBO And mSelectOpen _
                And mListHighlight >= 1 Then
                ' A keyboard pick commits exactly like a mouse pick,
                ' handlers included, and only when the value changes.
                If PickComboItem(hostApp, mListHighlight) Then
                    RunClickHandlers hostApp
                End If
                Exit Sub
            Else
                BlurField hostApp, True
                Exit Sub
            End If
        Case "{TAB}", "{CTRLENTER}"
            ' Tab commits like Enter: the form-field convention for
            ' leaving a field with its value kept.
            BlurField hostApp, True
            Exit Sub
        Case "{ESC}"
            ' A combo clears first, then leaves: the first Esc empties
            ' the value and reopens the unfiltered list with focus kept;
            ' a second Esc on the empty combo blurs and commits the
            ' clear, so state never disagrees with the face. A TextInput
            ' keeps the revert-on-Esc convention.
            If mKind = KIND_COMBO Then
                If LenB(mText) > 0 Then
                    mText = vbNullString
                    mValue = 0
                    mCaretPos = 0
                Else
                    BlurField hostApp, True
                    Exit Sub
                End If
            Else
                BlurField hostApp, False
                Exit Sub
            End If
        Case "{BS}"
            If mCaretPos > 0 Then
                mText = Left$(mText, mCaretPos - 1) & Mid$(mText, mCaretPos + 1)
                mCaretPos = mCaretPos - 1
            End If
        Case "{DEL}"
            If mCaretPos < Len(mText) Then
                mText = Left$(mText, mCaretPos) & Mid$(mText, mCaretPos + 2)
            End If
        Case "{LEFT}"
            If mCaretPos > 0 Then mCaretPos = mCaretPos - 1
        Case "{RIGHT}"
            If mCaretPos < Len(mText) Then mCaretPos = mCaretPos + 1
        Case "{UP}"
            If mMultiLine Then
                MoveCaretLine -1
            ElseIf mKind = KIND_COMBO Then
                ComboListNav hostApp, -1
            End If
        Case "{DOWN}"
            If mMultiLine Then
                MoveCaretLine 1
            ElseIf mKind = KIND_COMBO Then
                ComboListNav hostApp, 1
            End If
        Case "{HOME}"
            MoveCaretLineEdge False
        Case "{END}"
            MoveCaretLineEdge True
        Case Else
            InsertAtCaret keyText
    End Select
    mCaretOn = True
    mCaretAccumMs = 0
    ' Live filtering, the payoff cells could never give: an edit leaves
    ' the pristine full list and re-filters by the new text, rewinding
    ' the window and dropping the keyboard highlight. Keys that change
    ' nothing reopen a closed list as it was.
    If mKind = KIND_COMBO Then
        If mText <> textBeforeKey Then
            mSelectOpen = True
            mFilterPristine = False
            mListOffset = 0
            mListHighlight = 0
        ElseIf Not mSelectOpen Then
            OpenDropList hostApp.Sheet
        End If
    End If
    Touch
End Sub

' Up/Down walk the drop list: the highlight moves through the matches
' and the visible window follows it. Down on a closed list opens it.
Private Sub ComboListNav(ByVal hostApp As ReDimUI, ByVal rowDelta As Long)
    Dim matchCount As Long
    Dim currentItem As Long

    If Not mSelectOpen Then OpenDropList hostApp.Sheet
    Set mListMatches = BuildListMatches(hostApp.Sheet)
    matchCount = mListMatches.Count
    If matchCount = 0 Then Exit Sub
    If mListHighlight = 0 Then
        ' A pristine list walks on from the item the text names, which
        ' is also its row; otherwise the walk starts at the window's
        ' first visible row (after mouse paging, or on a first press).
        If mFilterPristine Then
            currentItem = ExactItemMatch(FieldText(hostApp.Sheet))
        End If
        If currentItem > 0 Then
            mListHighlight = currentItem + rowDelta
        Else
            mListHighlight = mListOffset + 1
        End If
    Else
        mListHighlight = mListHighlight + rowDelta
    End If
    If mListHighlight < 1 Then mListHighlight = 1
    If mListHighlight > matchCount Then mListHighlight = matchCount
    If mListHighlight > mListOffset + mListRows Then
        mListOffset = mListHighlight - mListRows
    ElseIf mListHighlight <= mListOffset Then
        mListOffset = mListHighlight - 1
    End If
End Sub

Private Sub InsertAtCaret(ByVal insertText As String)
    mText = Left$(mText, mCaretPos) & insertText & Mid$(mText, mCaretPos + 1)
    mCaretPos = mCaretPos + Len(insertText)
End Sub

' The caret's hard-line index and column, walking newline characters up
' to the caret. Vertical moves work in hard lines: a wrapped visual line
' counts as one, which keeps the model honest and the math exact.
Private Sub CaretLineCol(ByRef lineIndex As Long, ByRef lineStart As Long, _
    ByRef caretColumn As Long)
    Dim probe As Long

    lineIndex = 0
    lineStart = 0
    For probe = 1 To mCaretPos
        If Mid$(mText, probe, 1) = vbLf Then
            lineIndex = lineIndex + 1
            lineStart = probe
        End If
    Next probe
    caretColumn = mCaretPos - lineStart
End Sub

Private Sub MoveCaretLine(ByVal lineDelta As Long)
    Dim hardLines() As String
    Dim lineIndex As Long
    Dim lineStart As Long
    Dim caretColumn As Long
    Dim targetLine As Long
    Dim newPos As Long
    Dim walk As Long

    hardLines = Split(mText, vbLf)
    CaretLineCol lineIndex, lineStart, caretColumn
    targetLine = lineIndex + lineDelta
    If targetLine < 0 Then targetLine = 0
    If targetLine > UBound(hardLines) Then targetLine = UBound(hardLines)
    If targetLine = lineIndex Then Exit Sub
    If caretColumn > Len(hardLines(targetLine)) Then
        caretColumn = Len(hardLines(targetLine))
    End If
    newPos = 0
    For walk = 0 To targetLine - 1
        newPos = newPos + Len(hardLines(walk)) + 1
    Next walk
    mCaretPos = newPos + caretColumn
End Sub

Private Sub MoveCaretLineEdge(ByVal toEnd As Boolean)
    Dim hardLines() As String
    Dim lineIndex As Long
    Dim lineStart As Long
    Dim caretColumn As Long
    Dim newPos As Long
    Dim walk As Long

    hardLines = Split(mText, vbLf)
    CaretLineCol lineIndex, lineStart, caretColumn
    newPos = 0
    For walk = 0 To lineIndex - 1
        newPos = newPos + Len(hardLines(walk)) + 1
    Next walk
    If toEnd Then newPos = newPos + Len(hardLines(lineIndex))
    mCaretPos = newPos
End Sub

' Outside-press blur, watched by the same pump frames as slider drags.
' Two independent signals, because neither alone covers everything:
' - The selection moved since focus: a cell was clicked or arrowed to.
'   The press itself can be invisible to polling (the grid's selection
'   mouse loop holds WM_TIMER until the button is back up), but the
'   selection it leaves behind is durable state the next frame sees.
'   Event-free, so it works even when a host disables application
'   events, and it runs in proven tick context.
' - A press edge lands off the field: shapes, chrome, other windows -
'   clicks that never touch the selection.
Friend Sub TickFieldWatch( _
    ByVal hostApp As ReDimUI, _
    ByVal buttonDown As Boolean, _
    ByVal pressEdge As Boolean, _
    ByVal pollSelection As Boolean _
)
    If Not IsFloatField() Then Exit Sub
    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    ' The selection compare costs three COM reads, so it rides the work
    ' cadence; a 50 ms commit latency is imperceptible. The press edge
    ' stays at frame rate.
    If pollSelection Then
        If CurrentSelectionKey() <> mFieldSelAtFocus Then
            BlurField hostApp, True
            Exit Sub
        End If
    End If
    If Not pressEdge Then Exit Sub
    If CursorOverField() Then Exit Sub
    BlurField hostApp, True
End Sub

Private Function CursorOverField() As Boolean
    Dim screenPoint As RdxScreenPoint
    Dim pointsX As Double
    Dim pointsY As Double
    Dim listDepth As Double

    If GetCursorPos(screenPoint) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(screenPoint.pixelX)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(screenPoint.pixelY)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    ' The open option list counts as part of the field.
    listDepth = DropListDepth()
    If pointsX < mLeft - 2 Or pointsX > mLeft + mWidth + 2 Then Exit Function
    If pointsY < mTop - 2 Or pointsY > mTop + mHeight + listDepth + 2 Then
        Exit Function
    End If
    CursorOverField = True
End Function

' Item indexes whose text contains the filter, case-insensitively; every
' index when the filter is empty.
Private Function FilterItems(ByVal filterText As String) As Collection
    Dim idx As Long

    Set FilterItems = New Collection
    If mItems Is Nothing Then Exit Function
    For idx = 1 To mItems.Count
        If LenB(filterText) = 0 Then
            FilterItems.Add idx
        ElseIf InStr(1, CStr(mItems.Item(idx)), filterText, _
            vbTextCompare) > 0 Then
            FilterItems.Add idx
        End If
    Next idx
End Function

Private Function ExactItemMatch(ByVal candidate As String) As Long
    Dim idx As Long

    If mItems Is Nothing Then Exit Function
    For idx = 1 To mItems.Count
        If StrComp(CStr(mItems.Item(idx)), candidate, _
            vbTextCompare) = 0 Then
            ExactItemMatch = idx
            Exit Function
        End If
    Next idx
End Function

' An option pick writes the text (a cell-backed combo's cell without
' re-entering SheetChange), takes the item as the selection, and closes
' the list. Like the field's other commit paths it writes WritesTo state
' only when the value changed, and it returns whether it did, so callers
' run handlers for real changes only. The value a pick replaces is what a
' focused float field held when focus arrived, since text typed to filter
' the list was never a value; otherwise it is the field's current text.
Friend Function PickComboItem( _
    ByVal hostApp As ReDimUI, _
    ByVal matchRow As Long _
) As Boolean
    Dim itemIndex As Long
    Dim priorEvents As Boolean
    Dim picked As String
    Dim replacedText As String
    Dim cell As Range

    If mListMatches Is Nothing Then Exit Function
    If matchRow < 1 Or matchRow > mListMatches.Count Then Exit Function
    itemIndex = CLng(mListMatches.Item(matchRow))
    picked = ItemAt(itemIndex)
    replacedText = FieldText(hostApp.Sheet)
    If IsFloatField() Then
        If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
            replacedText = mFieldTextAtFocus
            ReDimUI.ClearKeyboardFocus
            mCaretOn = False
        End If
        mText = picked
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
    PickComboItem = (picked <> replacedText)
    If PickComboItem And LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, picked
    End If
    Touch
End Function

Private Sub ApplyInputCell( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim cell As Range

    If Not mAppliedValid Or mAppliedCellFill <> themeValue.SurfaceColor Then
        Set cell = InputCell(hostSheet)
        cell.Font.Name = themeValue.FontName
        cell.Font.Size = themeValue.BaseFontSize
        cell.Font.Color = themeValue.OnSurfaceColor
        cell.Interior.Color = themeValue.SurfaceColor
        cell.HorizontalAlignment = xlLeft
        ' Unlocked from birth so ProtectSurface keeps inputs editable.
        cell.Locked = False
        mAppliedCellFill = themeValue.SurfaceColor
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

Friend Sub DeleteOwnShapes(ByVal hostSheet As Worksheet)
    Dim idx As Long
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
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "optu").Delete
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "optd").Delete
    For idx = 1 To optionCount
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "c" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "d" & idx).Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & idx).Delete
    Next idx
    On Error GoTo 0
End Sub

Friend Function NeedsAnimation() As Boolean
    NeedsAnimation = (mKind = KIND_SPINNER And mVisible) _
        Or (mKind = KIND_TOAST And (mSlideRemaining > 0 Or mToastExiting)) _
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
            ' Re-pin the frame after every rotation write. Probing never
            ' caught rotation moving the frame at the object-model level,
            ' but a spinner was observed walking down-right during long
            ' multi-feed runs in interactive use; if the frame ever
            ' deviates, this snaps it back the same frame, and the
            ' corrective write forces a repaint at the true rect.
            With hostSheet.Shapes(ShapeName)
                .IncrementRotation elapsedMs * 0.6
                If Abs(.Left - mLeft) > 0.01 Then .Left = mLeft
                If Abs(.Top - mTop) > 0.01 Then .Top = mTop
            End With
        Case KIND_TOAST
            mSlideRemaining = mSlideRemaining * (0.55 ^ (elapsedMs / 50#))
            If mSlideRemaining < 0.75 Then mSlideRemaining = 0
            hostSheet.Shapes(ShapeName).Top = mTop + mSlideRemaining
            If mToastExiting And mExitRemaining > 0 Then
                mExitRemaining = mExitRemaining - elapsedMs / 180#
                If mExitRemaining < 0 Then mExitRemaining = 0
                With hostSheet.Shapes(ShapeName)
                    .Fill.Transparency = 1 - mExitRemaining
                    .Line.Transparency = 1 - mExitRemaining
                    .TextFrame2.TextRange.Font.Fill.Transparency = _
                        1 - mExitRemaining
                End With
            End If
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
Friend Sub ConfigureClickOp( _
    ByVal handlerProc As String, _
    ByVal targetComponentId As String _
)
    RequireRole ROLE_ASYNC_OP, "ConfigureClickOp"
    mOpRunsProc = handlerProc
    Set mOpTask = Nothing
    Set mOpDisableIds = New Collection
    mOpDisableIds.Add targetComponentId
End Sub

Public Function RunsProc(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "RunsProc"
    RequireNotRunning "RunsProc"
    mOpRunsProc = handlerProc
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
    Dim idx As Long

    RequireRole ROLE_ASYNC_OP, "Disables"
    RequireNotRunning "Disables"
    For idx = LBound(componentIds) To UBound(componentIds)
        mOpDisableIds.Add CStr(componentIds(idx))
    Next idx
    Set Disables = Me
End Function

Public Function ShowsSpinner(ByVal spinnerId As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "ShowsSpinner"
    RequireNotRunning "ShowsSpinner"
    mOpSpinnerId = spinnerId
    Set ShowsSpinner = Me
End Function

Public Function OnDone(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnDone"
    mOpDoneProc = handlerProc
    Set OnDone = Me
End Function

Public Function OnFail(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnFail"
    mOpFailProc = handlerProc
    Set OnFail = Me
End Function

Public Function OnCancel(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_ASYNC_OP, "OnCancel"
    mOpCancelProc = handlerProc
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

Friend Sub TickOp(ByVal hostApp As ReDimUI)
    If Not mOpRunning Then Exit Sub
    mOpTask.AdvanceTask
    If Not mOpTask.IsCompleted Then Exit Sub
    FinishOp hostApp
End Sub

Private Sub FinishOp(ByVal hostApp As ReDimUI)
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
    SetOpControlsBusy hostApp, False
    If LenB(mOpStatusKey) > 0 Then hostApp.SetState mOpStatusKey, outcomeState
    If mOpTask.IsFaulted Then
        hostApp.ReportAppError "ReDim async op '" & mComponentId & _
            "' failed: " & mOpLastError
    End If
    If LenB(outcomeProc) > 0 Then
        On Error Resume Next
        Application.Run outcomeProc
        On Error GoTo 0
    End If
End Sub

Private Sub SetOpControlsBusy(ByVal hostApp As ReDimUI, ByVal busyValue As Boolean)
    Dim disabledId As Variant
    Dim componentValue As ReDimUI
    Dim spinnerValue As ReDimUI

    For Each disabledId In mOpDisableIds
        If hostApp.TryGetComponent(CStr(disabledId), componentValue) Then
            componentValue.SetBusy busyValue
        End If
    Next disabledId
    If LenB(mOpSpinnerId) > 0 Then
        If hostApp.TryGetComponent(mOpSpinnerId, spinnerValue) Then
            spinnerValue.Visible busyValue
        End If
    End If
End Sub

Private Sub RequireNotRunning(ByVal memberName As String)
    If mOpRunning Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            "Async op '" & mComponentId & "' is already running."
    End If
    If mJobRunning Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
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
Public Function Steps(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_JOB, "Steps"
    RequireNotRunning "Steps"
    mJobStepProc = handlerProc
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
Public Function PacedMs(ByVal paceMs As Long) As ReDimUI
    RequireRole ROLE_JOB, "PacedMs"
    If paceMs < 1 Then paceMs = 1
    mJobPaceMs = paceMs
    Set PacedMs = Me
End Function

Public Function JobOnDone(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnDone"
    mJobDoneProc = handlerProc
    Set JobOnDone = Me
End Function

Public Function JobOnFail(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnFail"
    mJobFailProc = handlerProc
    Set JobOnFail = Me
End Function

Public Function JobOnCancel(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_JOB, "JobOnCancel"
    mJobCancelProc = handlerProc
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
    ByVal hostApp As ReDimUI, _
    ByVal nowMs As LongLong, _
    ByVal runWork As Boolean, _
    ByVal animating As Boolean _
)
    Dim stepDone As Variant
    Dim effectiveBudget As Long

    If Not mJobRunning Then Exit Sub
    If mJobCancelRequested Then
        FinishJob hostApp, mJobCancelProc
        Exit Sub
    End If
    On Error GoTo StepFailed
    If mJobPaceMs > 0 Then
        If nowMs - mJobLastStepMs < CLngLng(mJobPaceMs) Then Exit Sub
        mJobLastStepMs = nowMs
        stepDone = Application.Run(mJobStepProc)
        If CBool(stepDone) Then FinishJob hostApp, mJobDoneProc
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
            FinishJob hostApp, mJobDoneProc
            Exit Sub
        End If
        If mJobCancelRequested Then
            FinishJob hostApp, mJobCancelProc
            Exit Sub
        End If
    Loop While PreciseNowMs() - startedPreciseMs < effectiveBudget
    On Error GoTo 0
    Exit Sub

StepFailed:
    Dim failure As String
    failure = "ReDim job '" & mComponentId & "' failed: " & Err.Description
    On Error GoTo 0
    FinishJob hostApp, mJobFailProc
    hostApp.ReportAppError failure
End Sub

Private Sub FinishJob(ByVal hostApp As ReDimUI, ByVal outcomeProc As String)
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
    ByVal primaryRgb As Long, ByVal onPrimaryRgb As Long, _
    ByVal surfaceRgb As Long, ByVal onSurfaceRgb As Long, _
    ByVal mutedRgb As Long, ByVal onMutedRgb As Long, _
    ByVal successRgb As Long, ByVal dangerRgb As Long, _
    ByVal borderRgb As Long, ByVal canvasRgb As Long, _
    ByVal fontFace As String, ByVal fontPoints As Double _
)
    mRole = ROLE_THEME
    mThemePrimary = primaryRgb
    mThemeOnPrimary = onPrimaryRgb
    mThemeSurface = surfaceRgb
    mThemeOnSurface = onSurfaceRgb
    mThemeMuted = mutedRgb
    mThemeOnMuted = onMutedRgb
    mThemeSuccess = successRgb
    mThemeDanger = dangerRgb
    mThemeBorder = borderRgb
    mThemeCanvas = canvasRgb
    mThemeFontName = fontFace
    mThemeFontSize = fontPoints
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

Public Function WithPrimary( _
    ByVal primaryRgb As Long, _
    ByVal onPrimaryRgb As Long _
) As ReDimUI
    RequireRole ROLE_THEME, "WithPrimary"
    mThemePrimary = primaryRgb
    mThemeOnPrimary = onPrimaryRgb
    Set WithPrimary = Me
End Function

Public Function WithFont( _
    ByVal fontFace As String, _
    ByVal fontPoints As Double _
) As ReDimUI
    RequireRole ROLE_THEME, "WithFont"
    mThemeFontName = fontFace
    mThemeFontSize = fontPoints
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
    ByVal fullShapeName As String, _
    ByRef targetAppId As String, _
    ByRef targetComponentId As String, _
    ByRef targetPart As String _
) As Boolean
    Dim remainder As String
    Dim separatorAt As Long
    Dim partAt As Long

    targetPart = vbNullString
    If Left$(fullShapeName, Len(SHAPE_PREFIX)) <> SHAPE_PREFIX Then Exit Function
    remainder = Mid$(fullShapeName, Len(SHAPE_PREFIX) + 1)
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
    ByVal memberName As String, _
    ByVal argumentName As String _
)
    Dim idx As Long
    Dim code As Integer

    If LenB(candidate) = 0 Then
        RaiseError ERROR_INVALID_ARGUMENT, memberName, argumentName & " is required."
    End If
    For idx = 1 To Len(candidate)
        code = Asc(UCase$(Mid$(candidate, idx, 1)))
        If Not ((code >= 65 And code <= 90) Or (code >= 48 And code <= 57) _
            Or (allowUnderscore And code = 95)) Then
            RaiseError ERROR_INVALID_ARGUMENT, memberName, _
                argumentName & " may contain only letters, digits" & _
                IIf(allowUnderscore, ", and single underscores.", ".")
        End If
    Next idx
    If allowUnderscore Then
        If InStr(1, candidate, PART_SEPARATOR) > 0 Then
            RaiseError ERROR_INVALID_ARGUMENT, memberName, _
                argumentName & " may not contain a double underscore."
        End If
    End If
End Sub

Private Sub RequireRole(ByVal expectedRole As Long, ByVal memberName As String)
    If mRole <> expectedRole Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            memberName & " is not valid on this ReDimUI value."
    End If
End Sub

Private Sub RaiseError( _
    ByVal errNumber As Long, _
    ByVal memberName As String, _
    ByVal messageText As String _
)
    Err.Raise errNumber, "ReDimUI." & memberName, messageText
End Sub

Private Function BoolToMso(ByVal flag As Boolean) As Long
    If flag Then BoolToMso = msoTrue Else BoolToMso = msoFalse
End Function

' Millisecond clock from QueryPerformanceCounter: sub-millisecond precision
' regardless of the system timer resolution. Deltas only; its epoch is not
' comparable to GetTickCount64.
Private Function PreciseNowMs() As Double
    Static ticksPerSecond As LongLong
    Dim counter As LongLong

    If ticksPerSecond = 0 Then QueryPerformanceFrequency ticksPerSecond
    QueryPerformanceCounter counter
    PreciseNowMs = counter * 1000# / ticksPerSecond
End Function

Private Sub AssignVariant(ByRef receiver As Variant, ByVal payload As Variant)
    If IsObject(payload) Then
        Set receiver = payload
    Else
        receiver = payload
    End If
End Sub

Private Function TryGetItem( _
    ByVal store As Collection, _
    ByVal lookupKey As String, _
    ByRef found As Variant _
) As Boolean
    On Error GoTo Missing
    AssignVariant found, store.Item(lookupKey)
    TryGetItem = True
Missing:
End Function

Private Function TryGetObject( _
    ByVal store As Collection, _
    ByVal lookupKey As String, _
    ByRef found As ReDimUI _
) As Boolean
    On Error GoTo Missing
    Set found = store.Item(lookupKey)
    TryGetObject = True
Missing:
End Function

Private Sub ReplaceItem( _
    ByVal store As Collection, _
    ByVal lookupKey As String, _
    ByVal payload As Variant _
)
    On Error Resume Next
    store.Remove lookupKey
    On Error GoTo 0
    store.Add payload, lookupKey
End Sub
