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
Private Declare PtrSafe Function GetCaretBlinkTime Lib "user32" () As Long
Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoW" (ByVal actionCode As Long, _
    ByVal actionParam As Long, ByRef resultValue As Long, _
    ByVal profileFlags As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" ( _
    ByVal ownerWindow As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" ( _
    ByVal formatCode As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" ( _
    ByVal formatCode As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" ( _
    ByVal formatCode As Long, ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" ( _
    ByVal allocFlags As Long, ByVal byteCount As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" ( _
    ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" ( _
    ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" ( _
    ByVal memoryHandle As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal textAddress As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal toAddress As LongPtr, ByVal fromAddress As LongPtr, _
    ByVal byteCount As LongPtr)

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const VK_LEFT_BUTTON As Long = 1
Private Const VK_RIGHT_BUTTON As Long = 2
Private Const SPI_GETCLIENTAREAANIMATION As Long = &H1042
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE_ZEROED As Long = &H42

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

' List rows open with a gutter for the check that marks the current or
' selected item: the glyph, a tab, then the item, aligned by a tab stop
' this many points in.
Private Const LIST_GUTTER As Double = 14

' Transfer paging arrows are square pointer targets of this many points,
' 24 pixels at 96 DPI: the WCAG 2.5.8 minimum target size.
Private Const PAGER_TARGET As Double = 18

' Sentinel for "the source could not be loaded"; contains characters no
' real path can, so it never collides with one.
Private Const IMAGE_MISSING As String = "?missing?"

Private Const VARIANT_PRIMARY As Long = 0
Private Const VARIANT_SECONDARY As Long = 1
Private Const VARIANT_SUCCESS As Long = 2
Private Const VARIANT_DANGER As Long = 3
Private Const VARIANT_NEUTRAL As Long = 4
Private Const VARIANT_WARNING As Long = 5

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
' Set while focus moves from one field to another, so the outgoing
' field's blur leaves the OnKey bindings in place for the incoming one.
Private mFocusHandoff As Boolean
' Test seam: the next real tick treats the field-watch press edge as
' fired, so in-tick blur paths are testable without physical input.
Private mForcePressEdge As Boolean
' Test seam: a fixed pointer position, in sheet points, that stands in
' for the cursor until cleared.
' The Alt+letter chords bound for the active sheet's access keys.
Private mBoundAccessKeys As String
Private mPointerOverridden As Boolean
Private mPointerX As Double
Private mPointerY As Double
' The caret blink interval Windows reported when a field last took
' focus; 0 when the user turned blinking off.
Private mCaretBlinkMs As Long
' Motion preference: follow Windows (0), reduced (1), or full (2).
Private mMotionChoice As Long
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
' The components a pump frame can have work for (spinners, toasts,
' sliders, text fields), so frames skip the rest; rebuilt lazily after
' components come or go.
Private mTickables As Collection
Private mTickablesStale As Boolean
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
' The button Enter clicks when focus is on a control that does not use
' Enter itself, and the control that held focus when a modal opened.
Private mDefaultButtonId As String
Private mFocusBeforeModal As String
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
' Set once a variant builder ran; a toast reads its tone only then.
Private mVariantChosen As Boolean
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
Private mToastActionLabel As String
Private mToastActionProc As String
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
' Composite part signatures: the kind's extra parts (tick box, knob,
' radio rows, stepper, slider, drop-list caret), drop-list rows, pagers,
' and empty row, check-list rows, transfer headers, rows, and scrollers.
' A part is rewritten only when its signature changes. Render forgets
' the extras signature; a recreated main shape forgets them all.
Private mAppliedExtrasSig As String
Private mAppliedRowStates() As String
Private mAppliedPagerUpSig As String
Private mAppliedPagerDownSig As String
Private mAppliedEmptyRowSig As String
Private mAppliedCheckHeaderSig As String
Private mAppliedHeaderLeftSig As String
Private mAppliedHeaderRightSig As String
Private mAppliedLeftRowStates() As String
Private mAppliedRightRowStates() As String
Private mAppliedScrollLeftSig As String
Private mAppliedScrollRightSig As String
Private mAppliedCellFill As Long
' AltText's override of the generated description, and the description
' last written to the shape.
Private mAltText As String
Private mAppliedAltText As String
' The caret index: characters before the caret, 0 through Len(mText).
Private mCaretPos As Long
' The selection runs from this anchor to the caret; equal means none.
Private mSelAnchor As Long
' Undo and redo snapshots (text, caret, anchor) for the focus session,
' and the kind and time of the last edit, which group typing.
Private mUndoStack As Collection
Private mRedoStack As Collection
Private mLastEditKind As String
Private mLastEditMs As LongLong
' The focused face's characters mapped to the text characters they show
' (0 for the insertion bar, ellipses, and completion letters).
Private mFaceMap() As Long
Private mFaceLength As Long
' Selection marks painted on the face now, and whether the face has ever
' carried a highlight, which Excel cannot take off again.
Private mMarksPainted As Boolean
Private mHighlightTouched As Boolean
' The last click in the field and where it put the caret, which tell a
' double click from two single ones.
Private mLastFieldClickMs As LongLong
Private mLastFieldClickPos As Long
' Field rules: the placeholder, the OnInput handler and its debounce with
' the time a pending call falls due, Numeric's allowances, MaxLength, and
' Validates' check with its current message; once a message shows, every
' edit rechecks until the text passes.
Private mPlaceholder As String
Private mInputProc As String
Private mDebounceMs As Long
Private mInputDueMs As LongLong
Private mNumeric As Boolean
Private mNumericDecimal As Boolean
Private mNumericNegative As Boolean
Private mMaxLength As Long
Private mValidateProc As String
Private mValidationError As String
Private mValidationLive As Boolean
Private mAppliedErrorBorder As Boolean
Private mAppliedMessageSig As String
Private mAppliedCountSig As String
' AutoGrow: the line cap (0 for off) and how far the field has grown past
' the height it was given; the geometry setters start that over.
Private mAutoGrowLines As Long
Private mGrowAdded As Double
' Combo assists: commits limited to the items, where the face shows the
' completion letters, and where they were painted last.
Private mRestrictToItems As Boolean
Private mGhostFaceStart As Long
Private mGhostFaceLength As Long
Private mAppliedGhostKey As String
' Windowed long lists: a SelectBox or ComboBox drop list shows at most
' mListRows of mListMatches (item indexes in list order) from
' mListOffset, with mListHighlight as the combo's keyboard cursor
' (1-based into the matches, 0 for none). A pristine combo list is
' unfiltered while its text names an item, until the user edits it.
' Transfer panels scroll through per-side offsets.
Private mListRows As Long
Private mListMatches As Collection
Private mListMatchesFilter As String
' Set when the list opened upward for want of room below the face.
Private mListFlipped As Boolean
' Items as a String array (1-based), rebuilt by one For Each after any
' change: a VBA Collection read by index walks the list, so loops that
' indexed mItems went quadratic on long lists.
Private mItemTexts() As String
Private mItemTextsCurrent As Boolean
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
' The grid selection an open list's watch compares against; empty until
' the watch's first frame records it.
Private mListSelAtOpen As String
' Tab order: 0 follows creation order, positive values come first in
' ascending order, and -1 leaves the control out of Tab (Focus still
' reaches it).
Private mTabIndex As Long
' The row a focused check list or transfer list's arrow keys have
' reached (1-based; 0 is a check list's select-all header), and for a
' transfer list the panel it is in (0 available, 1 chosen).
Private mKeyRow As Long
Private mKeyPanel As Long
Private mKeyCursorSet As Boolean
' Type-ahead in a focused select: the prefix typed so far and when its
' last letter landed; a second's pause starts a new prefix.
Private mTypeAhead As String
Private mTypeAheadMs As LongLong
Private mAppliedRingSig As String
Private mAppliedKeyCursorSig As String
' AccessKey's letter or digit (lower case) and the one last underlined.
Private mAccessKey As String
Private mAppliedAccessKey As String
' Clearable: a float field shows a clear button while it holds text.
Private mClearable As Boolean
Private mAppliedClearSig As String
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

' Access-key chords follow the sheet in front.
Private Sub mExcelApp_SheetActivate(ByVal Sh As Object)
    RefreshAccessKeys
End Sub

Private Sub mExcelApp_WorkbookActivate(ByVal Wb As Workbook)
    RefreshAccessKeys
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
    ' A click anywhere else closes an open list, even a click the target
    ' then ignores.
    appValue.CloseListsExcept targetComponentId
    ' It also ends keyboard focus held elsewhere, so a field commits
    ' before the clicked control's handler runs. A float field clicked
    ' to focus takes focus over itself, and a click on a modal's chrome
    ' leaves the dialog's focus alone.
    If LenB(mFocusComponentId) > 0 Then
        If mFocusAppId <> targetAppId _
            Or mFocusComponentId <> targetComponentId Then
            If Not componentValue.IsFloatField _
                And Left$(targetComponentId, 4) <> "mdl_" Then
                EndKeyboardFocus
            End If
        End If
    End If
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
                mFocusHandoff = True
                previousComponent.LeaveFocus previousApp, True
                mFocusHandoff = False
            End If
        End If
    End If
    mFocusAppId = hostApp.AppId
    mFocusComponentId = componentValue.ComponentId
    mCaretBlinkMs = SystemCaretBlinkMs()
    RdxBindKeys
    RequestPump
End Sub

' The Windows caret blink interval, or 0 when the user turned blinking
' off (GetCaretBlinkTime reports INFINITE) or the call fails.
Private Function SystemCaretBlinkMs() As Long
    Dim reportedMs As Long

    reportedMs = GetCaretBlinkTime()
    If reportedMs > 0 Then SystemCaretBlinkMs = reportedMs
End Function

Friend Property Get CaretBlinkMs() As Long
    If Not IsFactory(Me) Then
        CaretBlinkMs = ReDimUI.CaretBlinkMs
        Exit Property
    End If
    CaretBlinkMs = mCaretBlinkMs
End Property

' Runs a handler with ReDimUI.Sender set to the component, the way a
' click does, for handlers fired from keys or the pump; a failure goes
' to the app's error handler.
Friend Sub RunHandlerAs( _
    ByVal hostApp As ReDimUI, _
    ByVal componentValue As ReDimUI, _
    ByVal handlerProc As String _
)
    Dim priorApp As ReDimUI
    Dim priorComponent As ReDimUI
    Dim failure As String

    If Not IsFactory(Me) Then
        ReDimUI.RunHandlerAs hostApp, componentValue, handlerProc
        Exit Sub
    End If
    Set priorApp = mSenderApp
    Set priorComponent = mSenderComponent
    Set mSenderApp = hostApp
    Set mSenderComponent = componentValue
    On Error GoTo HandlerFailed
    Application.Run handlerProc
    On Error GoTo 0
    Set mSenderApp = priorApp
    Set mSenderComponent = priorComponent
    Exit Sub
HandlerFailed:
    failure = "ReDim handler error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
    Set mSenderApp = priorApp
    Set mSenderComponent = priorComponent
    hostApp.ReportAppError failure
End Sub

' Binds Alt plus each access key the active sheet's apps declare and
' releases the rest, so the chords belong to ReDim only while an app
' sheet is in front.
Friend Sub RefreshAccessKeys()
    Dim appValue As ReDimUI
    Dim wantedKeys As String
    Dim declaredKeys As String
    Dim idx As Long
    Dim keyLetter As String

    If Not IsFactory(Me) Then
        ReDimUI.RefreshAccessKeys
        Exit Sub
    End If
    EnsureFactory
    On Error Resume Next
    For Each appValue In mApps
        If appValue.Sheet Is ActiveSheet Then
            declaredKeys = appValue.AccessKeyLetters
            For idx = 1 To Len(declaredKeys)
                keyLetter = Mid$(declaredKeys, idx, 1)
                If InStr(wantedKeys, keyLetter) = 0 Then
                    wantedKeys = wantedKeys & keyLetter
                End If
            Next idx
        End If
    Next appValue
    For idx = 1 To Len(mBoundAccessKeys)
        keyLetter = Mid$(mBoundAccessKeys, idx, 1)
        If InStr(wantedKeys, keyLetter) = 0 Then Application.OnKey "%" & keyLetter
    Next idx
    For idx = 1 To Len(wantedKeys)
        keyLetter = Mid$(wantedKeys, idx, 1)
        If InStr(mBoundAccessKeys, keyLetter) = 0 Then
            Application.OnKey "%" & keyLetter, "'RdxAccessKey """ & keyLetter & """'"
        End If
    Next idx
    On Error GoTo 0
    mBoundAccessKeys = wantedKeys
End Sub

' Alt plus a letter: the first app on the active sheet with a control
' declaring it clicks that control.
Friend Sub DispatchAccessKey(ByVal keyLetter As String)
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.DispatchAccessKey keyLetter
        Exit Sub
    End If
    EnsureFactory
    For Each appValue In mApps
        If appValue.Sheet Is ActiveSheet Then
            If appValue.ActivateAccessKey(LCase$(keyLetter)) Then Exit Sub
        End If
    Next appValue
End Sub

' Ends keyboard focus the way a click elsewhere does: a field commits,
' any other control drops its ring.
Friend Sub EndKeyboardFocus()
    Dim focusApp As ReDimUI
    Dim focusedValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.EndKeyboardFocus
        Exit Sub
    End If
    If LenB(mFocusComponentId) = 0 Then Exit Sub
    Set focusApp = FindApp(mFocusAppId)
    If Not focusApp Is Nothing Then
        Set focusedValue = focusApp.PeekComponent(mFocusComponentId)
    End If
    If focusedValue Is Nothing Then
        ClearKeyboardFocus
    Else
        focusedValue.LeaveFocus focusApp, True
    End If
End Sub

Friend Sub ClearKeyboardFocus()
    If Not IsFactory(Me) Then
        ReDimUI.ClearKeyboardFocus
        Exit Sub
    End If
    mFocusAppId = vbNullString
    mFocusComponentId = vbNullString
    If Not mFocusHandoff Then RdxReleaseKeys
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

' Test seam: pointer-driven behavior (outside presses, slider drags) reads
' this sheet position instead of the cursor until ClearPointerOverride.
Friend Sub OverridePointer(ByVal pointsX As Double, ByVal pointsY As Double)
    If Not IsFactory(Me) Then
        ReDimUI.OverridePointer pointsX, pointsY
        Exit Sub
    End If
    mPointerOverridden = True
    mPointerX = pointsX
    mPointerY = pointsY
End Sub

Friend Sub ClearPointerOverride()
    If Not IsFactory(Me) Then
        ReDimUI.ClearPointerOverride
        Exit Sub
    End If
    mPointerOverridden = False
End Sub

Friend Function PointerOverride( _
    ByRef pointsX As Double, _
    ByRef pointsY As Double _
) As Boolean
    If Not IsFactory(Me) Then
        PointerOverride = ReDimUI.PointerOverride(pointsX, pointsY)
        Exit Function
    End If
    If Not mPointerOverridden Then Exit Function
    pointsX = mPointerX
    pointsY = mPointerY
    PointerOverride = True
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

' Toasts slide in, glide up as others leave, and fade out, unless motion
' is reduced: then they appear, move, and leave at once. Reduction
' follows the Windows "Show animations in Windows" setting; pass True
' or False to override it, or nothing to follow Windows again.
Public Sub ReduceMotion(Optional ByVal motionOff As Variant)
    If Not IsFactory(Me) Then
        ReDimUI.ReduceMotion motionOff
        Exit Sub
    End If
    If IsMissing(motionOff) Then
        mMotionChoice = 0
    ElseIf CBool(motionOff) Then
        mMotionChoice = 1
    Else
        mMotionChoice = 2
    End If
End Sub

' True when animations should not play: the ReduceMotion override when
' set, else the Windows setting, read each time since users change it.
Public Property Get MotionReduced() As Boolean
    Dim animationsOn As Long

    If Not IsFactory(Me) Then
        MotionReduced = ReDimUI.MotionReduced
        Exit Property
    End If
    Select Case mMotionChoice
        Case 1
            MotionReduced = True
        Case 2
            MotionReduced = False
        Case Else
            animationsOn = 1
            If SystemParametersInfo(SPI_GETCLIENTAREAANIMATION, 0, _
                animationsOn, 0) <> 0 Then
                MotionReduced = (animationsOn = 0)
            End If
    End Select
End Property

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
    ' Windows go with their apps; a rebuild would otherwise tab to them.
    Set mWindows = Nothing
    Set mNavStack = Nothing
    mActiveWindowId = vbNullString
    mPointerOverridden = False
    RefreshAccessKeys
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

' White and yellow on black with white edges. Every pairing the
' controls draw clears WCAG AA, most by a wide margin; ContrastReport
' lists them.
Public Function ThemeHighContrast() As ReDimUI
    Dim themeValue As ReDimUI

    Set themeValue = New ReDimUI
    themeValue.ConfigureTheme _
        RGB(255, 255, 0), RGB(0, 0, 0), _
        RGB(0, 0, 0), RGB(255, 255, 255), _
        RGB(38, 38, 38), RGB(200, 200, 200), _
        RGB(0, 255, 0), RGB(255, 128, 128), _
        RGB(255, 255, 255), RGB(0, 0, 0), _
        "Segoe UI", 11
    Set ThemeHighContrast = themeValue
End Function

' The WCAG 2 contrast ratio of two RGB colors, from 1 (none) to 21
' (black on white). Text needs 4.5 (3 when large); the edges and
' accents that mark controls and their states need 3.
Public Function ContrastRatio(ByVal foreRgb As Long, ByVal backRgb As Long) As Double
    Dim foreLight As Double
    Dim backLight As Double

    foreLight = RelativeLuminance(foreRgb)
    backLight = RelativeLuminance(backRgb)
    If foreLight < backLight Then
        ContrastRatio = (backLight + 0.05) / (foreLight + 0.05)
    Else
        ContrastRatio = (foreLight + 0.05) / (backLight + 0.05)
    End If
End Function

Private Function RelativeLuminance(ByVal colorRgb As Long) As Double
    RelativeLuminance = 0.2126 * LinearChannel(colorRgb And &HFF&) _
        + 0.7152 * LinearChannel((colorRgb \ &H100&) And &HFF&) _
        + 0.0722 * LinearChannel((colorRgb \ &H10000) And &HFF&)
End Function

Private Function LinearChannel(ByVal channelValue As Long) As Double
    Dim fraction As Double

    fraction = channelValue / 255#
    If fraction <= 0.04045 Then
        LinearChannel = fraction / 12.92
    Else
        LinearChannel = ((fraction + 0.055) / 1.055) ^ 2.4
    End If
End Function

' One line per color pairing the controls draw: its contrast ratio, the
' ratio it needs, and pass or fail. Text pairs need 4.5:1; field edges
' and the accent that marks states need 3:1. Ratios are truncated, so a
' listed ratio never rounds up past a threshold it misses.
Public Function ContrastReport() As String
    RequireRole ROLE_THEME, "ContrastReport"
    ContrastReport = _
        ContrastLine("Body text", "OnSurface on Surface", _
            mThemeOnSurface, mThemeSurface, 4.5) & vbLf & _
        ContrastLine("Canvas text", "OnSurface on Canvas", _
            mThemeOnSurface, mThemeCanvas, 4.5) & vbLf & _
        ContrastLine("Primary button", "OnPrimary on Primary", _
            mThemeOnPrimary, mThemePrimary, 4.5) & vbLf & _
        ContrastLine("Success button", "OnPrimary on Success", _
            mThemeOnPrimary, mThemeSuccess, 4.5) & vbLf & _
        ContrastLine("Danger button", "OnPrimary on Danger", _
            mThemeOnPrimary, mThemeDanger, 4.5) & vbLf & _
        ContrastLine("Muted fill text", "OnSurface on Muted", _
            mThemeOnSurface, mThemeMuted, 4.5) & vbLf & _
        ContrastLine("Secondary text", "OnMuted on Surface", _
            mThemeOnMuted, mThemeSurface, 4.5) & vbLf & _
        ContrastLine("Field edge", "Border on Surface", _
            mThemeBorder, mThemeSurface, 3) & vbLf & _
        ContrastLine("Accent", "Primary on Surface", _
            mThemePrimary, mThemeSurface, 3)
End Function

Private Function ContrastLine( _
    ByVal pairName As String, _
    ByVal pairColors As String, _
    ByVal foreRgb As Long, _
    ByVal backRgb As Long, _
    ByVal neededRatio As Double _
) As String
    Dim shownRatio As Double

    shownRatio = Int(ContrastRatio(foreRgb, backRgb) * 100) / 100
    ContrastLine = pairName & " (" & pairColors & "): " & _
        Format$(shownRatio, "0.00") & ":1, needs " & _
        Format$(neededRatio, "0.0") & ":1, " & _
        IIf(shownRatio >= neededRatio, "pass", "fail")
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

' Transient notification card with a close button. Expires on the pump
' after ttlMs, or when clicked. Without ttlMs it stays long enough to
' read: three seconds plus 60 ms a character, from four seconds to
' twelve. Primary, Success, Warning, and Danger give it an info,
' success, warning, or error tone: an icon and a matching edge.
Public Function Toast( _
    ByVal messageText As String, _
    Optional ByVal ttlMs As Long = 0 _
) As ReDimUI
    Dim toastValue As ReDimUI

    RequireRole ROLE_APP, "Toast"
    If ttlMs <= 0 Then
        ttlMs = 3000 + 60 * Len(messageText)
        If ttlMs < 4000 Then ttlMs = 4000
        If ttlMs > 12000 Then ttlMs = 12000
    End If
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
    ' The dialog takes the keys: OK holds focus, Tab stays among its
    ' buttons, Enter confirms, and Esc cancels.
    mFocusBeforeModal = vbNullString
    If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
        mFocusBeforeModal = ReDimUI.FocusedComponentId
    End If
    Component("mdl_ok").TakeFocus Me, False
End Sub

Public Sub CloseModal()
    Dim heldByDialog As Boolean
    Dim returnTo As ReDimUI

    RequireRole ROLE_APP, "CloseModal"
    heldByDialog = Left$(ReDimUI.FocusedComponentId, 4) = "mdl_" _
        And ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId)
    BeginUpdate
    HideIfExists "mdl_ov"
    HideIfExists "mdl_card"
    HideIfExists "mdl_ok"
    HideIfExists "mdl_cancel"
    EndUpdate
    ' Focus goes back to the control that held it before the dialog.
    If heldByDialog Then
        If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
            ReDimUI.ClearKeyboardFocus
        End If
        If LenB(mFocusBeforeModal) > 0 Then
            If TryGetComponent(mFocusBeforeModal, returnTo) Then
                If returnTo.IsFocusable Then returnTo.TakeFocus Me, False
            End If
        End If
    End If
    mFocusBeforeModal = vbNullString
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

' Components placed Below or RightOf one that changed size follow it, and
' the ones placed against those follow in turn.
Friend Sub ReapplyDependents(ByVal anchorId As String)
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "ReapplyDependents"
    For Each componentValue In mComponents
        If componentValue.PlacedRelativeTo = anchorId Then
            componentValue.Apply mSheet, mTheme, Me
            ReapplyDependents componentValue.ComponentId
        End If
    Next componentValue
End Sub

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
    mComponentOrder.Add targetComponentId, targetComponentId
    mTickablesStale = True
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
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "Render"
    For Each orderedId In mComponentOrder
        NotifyDirty CStr(orderedId)
    Next orderedId
    ' An explicit render rewrites the single-shape composites' parts (tick
    ' box, knob, radio rows, stepper, slider, drop-list caret), so a part
    ' moved or deleted by hand comes back. List composites keep the row
    ' state they diff against, as their whole-list signatures did before.
    For Each componentValue In mComponents
        componentValue.ForgetExtrasSignature
    Next componentValue
    mRendered = True
    FlushDirty
    PruneOrphanShapes
    ' Anything rendered that needs frames (a slider's press watch, a
    ' spinner, a toast) arms the pump right here; demand alone only keeps
    ' an already-armed pump alive.
    If HasAppWork() Then ReDimUI.RequestPump
    ReDimUI.RefreshAccessKeys
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
    RequireRole ROLE_APP, "Unmount"
    ClearHotKeys
    If ReDimUI.FocusedComponentId <> vbNullString Then
        If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
            ReDimUI.ClearKeyboardFocus
        End If
    End If
    If mSurfaceProtected Then ProtectSurface False
    ' One pass over the sheet finds every shape carrying this app's
    ' prefix, parts and leftovers alike, and one call deletes them.
    If deleteShapes Then
        DeleteShapesNamed mSheet, SHAPE_PREFIX & mAppId & "_", vbNullString
    End If
    Set mComponents = New Collection
    Set mComponentOrder = New Collection
    Set mDirtyIds = New Collection
    Set mTickables = Nothing
    ReDimUI.ForgetApp mAppId
    ReDimUI.RefreshAccessKeys
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
    For Each componentValue In Tickables()
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
    For Each componentValue In Tickables()
        componentValue.TickAnimation mSheet, elapsedMs
        componentValue.TickSlideDrag Me, leftButtonDown
        componentValue.TickListWatch pressEdge, runWork
        If runWork Then
            If componentValue.IsExpiredToast(nowMs) Then
                componentValue.BeginToastExit
            End If
        End If
        If componentValue.IsToastExitDone Then
            expiredIds.Add componentValue.ComponentId
        End If
    Next componentValue
    ' The control holding keyboard focus, tickable or not, watches for a
    ' press elsewhere or a moved selection.
    If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
        If TryGetComponent(ReDimUI.FocusedComponentId, componentValue) Then
            componentValue.TickFocusWatch Me, pressEdge, runWork
        End If
    End If
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
    ' Keyboard focus keeps the frames coming for its click-away watch.
    If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
        HasAppWork = True
        Exit Function
    End If
    For Each componentValue In Tickables()
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
        If componentValue.NeedsListWatch Then
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

Friend Property Get IsTickable() As Boolean
    Select Case mKind
        Case KIND_SPINNER, KIND_TOAST, KIND_SLIDEBAR, KIND_INPUT, KIND_COMBO, _
            KIND_SELECT
            IsTickable = True
    End Select
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
    mGrowAdded = 0
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
    mGrowAdded = 0
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

' A multi-line float TextInput grows with its text from the height it was
' given, a line at a time up to maxLines, and shrinks back as lines go;
' components placed Below it move with it. AutoGrow turns MultiLine on,
' and 0 turns growing off.
Public Function AutoGrow(Optional ByVal maxLines As Long = 6) As ReDimUI
    RequireRole ROLE_COMPONENT, "AutoGrow"
    If mKind <> KIND_INPUT Then
        RaiseError ERROR_INVALID_OPERATION, "AutoGrow", _
            "AutoGrow applies to TextInput only."
    End If
    If maxLines < 0 Then
        RaiseError ERROR_INVALID_ARGUMENT, "AutoGrow", _
            "maxLines must be 0 or more."
    End If
    mAutoGrowLines = maxLines
    If maxLines > 0 Then mMultiLine = True
    Touch
    Set AutoGrow = Me
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

' Amber, for actions that need care. On a toast it is the warning tone.
Public Function Warning() As ReDimUI
    Set Warning = SetVariant(VARIANT_WARNING)
End Function

Private Function SetVariant(ByVal variantCode As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "Variant"
    mStyleVariant = variantCode
    mVariantChosen = True
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

' The alternative text screen readers announce for the control. ReDim
' writes one from the kind, text, and state ("Save, button", "Agree,
' checkbox, checked", ", unavailable" when disabled); AltText replaces
' it, and an empty string restores the generated one.
' A toast's action button, such as Undo: a click dismisses the toast and
' runs handlerProc. A toast with an action stays four seconds longer.
Public Function Action(ByVal actionLabel As String, ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Action"
    If mKind <> KIND_TOAST Then
        RaiseError ERROR_INVALID_ARGUMENT, "Action", "Action applies to toasts."
    End If
    If LenB(mToastActionLabel) = 0 Then mExpiresAtMs = mExpiresAtMs + 4000
    mToastActionLabel = actionLabel
    mToastActionProc = handlerProc
    Touch
    Set Action = Me
End Function

' Tab order: positive values come first, ascending; the rest follow in
' creation order; -1 takes the control out of Tab while Focus still
' reaches it.
Public Function TabIndex(ByVal tabSlot As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "TabIndex"
    mTabIndex = tabSlot
    Set TabIndex = Me
End Function

' Gives the control keyboard focus and scrolls it into view: the way in
' from code, since a click focuses only text fields.
Public Function Focus() As ReDimUI
    RequireRole ROLE_COMPONENT, "Focus"
    If IsFocusable() Then
        TakeFocus OwnerApp, False
        ScrollIntoSight
    End If
    Set Focus = Me
End Function

' Alt plus this letter or digit clicks the control while its sheet is
' active, and the first matching character of its text is underlined.
' For buttons and tick boxes; an ampersand in their text stays literal.
Public Function AccessKey(ByVal keyLetter As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "AccessKey"
    If Len(keyLetter) > 1 Or Not LCase$(keyLetter & "a") Like "[a-z0-9]*" Then
        RaiseError ERROR_INVALID_ARGUMENT, "AccessKey", _
            "AccessKey takes one letter or digit."
    End If
    mAccessKey = LCase$(keyLetter)
    Touch
    ReDimUI.RefreshAccessKeys
    Set AccessKey = Me
End Function

Friend Property Get AccessKeyLetter() As String
    AccessKeyLetter = mAccessKey
End Property

' A float TextInput or ComboBox shows a clear button inside its right
' end while it holds text. A click on it empties the field and leaves
' it focused; the emptied value commits when focus leaves.
Public Function Clearable(Optional ByVal clearOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Clearable"
    mClearable = clearOn
    Touch
    Set Clearable = Me
End Function

' A float ComboBox that commits only its items: a commit takes the item
' the text names or begins, in the item's spelling, keeps an empty text,
' and otherwise goes back to the text the field had when focus arrived.
' A cell-backed combo's cell is Excel's to restrict, with data validation.
Public Function RestrictToItems(Optional ByVal restrictOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "RestrictToItems"
    If mKind <> KIND_COMBO Then
        RaiseError ERROR_INVALID_OPERATION, "RestrictToItems", _
            "RestrictToItems applies to ComboBox only."
    End If
    mRestrictToItems = restrictOn
    Set RestrictToItems = Me
End Function

' A float TextInput or ComboBox shows this hint in muted ink while it is
' empty, focused or not.
Public Function Placeholder(ByVal hintText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Placeholder"
    mPlaceholder = hintText
    Touch
    Set Placeholder = Me
End Function

' Runs while the user types in a float field: after each edit, or once
' typing pauses for DebounceMs. ReDimUI.Sender is the field, and its
' InputValue holds the text so far.
Public Function OnInput(ByVal handlerProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnInput"
    mInputProc = handlerProc
    Set OnInput = Me
End Function

Public Function DebounceMs(ByVal pauseMs As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "DebounceMs"
    If pauseMs < 0 Then pauseMs = 0
    mDebounceMs = pauseMs
    Set DebounceMs = Me
End Function

' Digits only, plus one decimal separator (the locale's; . and , both
' type it) and a leading minus, unless turned off. Pasted text keeps
' what fits a number.
Public Function Numeric( _
    Optional ByVal allowDecimal As Boolean = True, _
    Optional ByVal allowNegative As Boolean = True _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "Numeric"
    mNumeric = True
    mNumericDecimal = allowDecimal
    mNumericNegative = allowNegative
    Set Numeric = Me
End Function

' Caps the text at maxChars: typing and pasting stop there, and a count
' under the field's right edge shows the length against the cap.
Public Function MaxLength(ByVal maxChars As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "MaxLength"
    If maxChars < 0 Then maxChars = 0
    mMaxLength = maxChars
    Touch
    Set MaxLength = Me
End Function

' Checks the text on commit with a Public Function that takes the text
' and returns "" when it is valid or a message when it is not. An
' invalid field shows a danger border and the message under it, and
' rechecks on every edit until the text passes. The commit still
' happens; ValidationError reads the message.
Public Function Validates(ByVal checkProc As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Validates"
    mValidateProc = checkProc
    Set Validates = Me
End Function

Public Property Get ValidationError() As String
    RequireRole ROLE_COMPONENT, "ValidationError"
    ValidationError = mValidationError
End Property

Public Property Get IsVisible() As Boolean
    RequireRole ROLE_COMPONENT, "IsVisible"
    IsVisible = mVisible
End Property

Public Function AltText(ByVal altDescription As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "AltText"
    mAltText = altDescription
    Touch
    Set AltText = Me
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
    ItemsChanged
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
    ItemsChanged
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
    Dim wantedSet As Collection
    Dim ignored As Variant

    RequireRole ROLE_COMPONENT, "CheckedFrom"
    ' Wanted texts go into a keyed set (Collection keys compare without
    ' case, like the match this replaced), so each item costs one lookup
    ' instead of a scan of the wanted list.
    Set wantedSet = New Collection
    For Each entry In ReadItemSource(itemSource, "CheckedFrom")
        On Error Resume Next
        wantedSet.Add True, CStr(entry)
        Err.Clear
        On Error GoTo 0
    Next entry
    Set mCheckedSet = New Collection
    If Not mItems Is Nothing Then
        EnsureItemTexts
        For idx = 1 To mItems.Count
            If TryGetItem(wantedSet, mItemTexts(idx), ignored) Then
                mCheckedSet.Add idx
            End If
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
    Dim area As Range
    Dim cellValues As Variant
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim idx As Long

    Set gathered = New Collection
    If TypeName(itemSource) = "Range" Then
        ' One Value read per area instead of one per cell, in the row-major
        ' order For Each over the cells gives.
        For Each area In itemSource.Areas
            If area.CountLarge = 1 Then
                If LenB(CStr(area.Value)) > 0 Then gathered.Add CStr(area.Value)
            Else
                cellValues = area.Value
                For rowIdx = 1 To UBound(cellValues, 1)
                    For colIdx = 1 To UBound(cellValues, 2)
                        If LenB(CStr(cellValues(rowIdx, colIdx))) > 0 Then
                            gathered.Add CStr(cellValues(rowIdx, colIdx))
                        End If
                    Next colIdx
                Next rowIdx
            End If
        Next area
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
    ItemsChanged
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
        EnsureItemTexts
        For probe = 1 To mItems.Count
            If mItemTexts(probe) = CStr(indexOrText) Then
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
    ItemsChanged
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
    ItemsChanged
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
            If Not mSelectOpen Then NoteListOpened
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
    If ReDimUI.MotionReduced Then mSlideRemaining = 0
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
    If ReDimUI.MotionReduced Then mExitRemaining = 0
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
    If ReDimUI.MotionReduced Then
        ' No glide: the next apply moves the shape to its slot, and its
        ' buttons with it.
        mSlideRemaining = 0
        Touch
    End If
    ReDimUI.RequestPump
End Sub

Friend Function IsExpiredToast(ByVal nowMs As LongLong) As Boolean
    If mKind <> KIND_TOAST Then Exit Function
    IsExpiredToast = (nowMs >= mExpiresAtMs)
End Function

' Test seam: milliseconds until the toast expires.
Friend Property Get ToastRemainingMs() As Long
    If mKind = KIND_TOAST Then ToastRemainingMs = CLng(mExpiresAtMs - GetTickCount64())
End Property

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
    If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
        ReDimUI.ClearKeyboardFocus
    End If
    DeleteOwnShapes hostApp.Sheet
    hostApp.ForgetComponent mComponentId
End Sub

Friend Sub ForgetComponent(ByVal targetComponentId As String)
    RequireRole ROLE_APP, "ForgetComponent"
    On Error Resume Next
    mComponents.Remove targetComponentId
    mComponentOrder.Remove targetComponentId
    Err.Clear
    On Error GoTo 0
    mTickablesStale = True
End Sub

Private Function Tickables() As Collection
    Dim componentValue As ReDimUI

    If mTickables Is Nothing Or mTickablesStale Then
        Set mTickables = New Collection
        For Each componentValue In mComponents
            If componentValue.IsTickable Then mTickables.Add componentValue
        Next componentValue
        mTickablesStale = False
    End If
    Set Tickables = mTickables
End Function

Friend Function IsBoundTo(ByVal stateKey As String) As Boolean
    IsBoundTo = (mBindTextKey = stateKey) Or (mBindValueKey = stateKey) _
        Or (mBindVisibleKey = stateKey) Or (mBindEnabledKey = stateKey) _
        Or (LenB(mBindSourceKey) > 0 And mBindSourceKey = stateKey)
End Function

' =====================================================================
' Transfer list moves
' =====================================================================

Private Sub EnsureTransferCollections()
    If mItems Is Nothing Then
        Set mItems = New Collection
        ItemsChanged
    End If
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

' The controls Tab visits: positive TabIndex values first, ascending,
' then the rest in creation order. While a modal is up only its buttons
' take part, which keeps focus inside the dialog.
Private Function TabOrder() As Collection
    Dim leading As Collection
    Dim trailing As Collection
    Dim orderedId As Variant
    Dim componentValue As ReDimUI
    Dim placed As ReDimUI
    Dim idx As Long
    Dim trapped As Boolean
    Dim inserted As Boolean

    Set leading = New Collection
    Set trailing = New Collection
    trapped = ModalUp()
    For Each orderedId In mComponentOrder
        If TryGetComponent(CStr(orderedId), componentValue) Then
            If componentValue.IsTabStop Then
                If Not trapped Or Left$(componentValue.ComponentId, 4) = "mdl_" Then
                    If componentValue.TabPosition > 0 Then
                        inserted = False
                        For idx = 1 To leading.Count
                            Set placed = leading.Item(idx)
                            If placed.TabPosition > componentValue.TabPosition Then
                                leading.Add componentValue, Before:=idx
                                inserted = True
                                Exit For
                            End If
                        Next idx
                        If Not inserted Then leading.Add componentValue
                    Else
                        trailing.Add componentValue
                    End If
                End If
            End If
        End If
    Next orderedId
    For idx = 1 To trailing.Count
        leading.Add trailing.Item(idx)
    Next idx
    Set TabOrder = leading
End Function

' Moves keyboard focus to the next or previous control in Tab order,
' wrapping, and scrolls it into view. False when no other control can
' take focus, so the caller can simply leave.
Friend Function MoveFocusFrom(ByVal currentId As String, ByVal goingForward As Boolean) As Boolean
    Dim stops As Collection
    Dim idx As Long
    Dim here As Long
    Dim nextStop As ReDimUI

    RequireRole ROLE_APP, "MoveFocusFrom"
    Set stops = TabOrder()
    If stops.Count = 0 Then Exit Function
    For idx = 1 To stops.Count
        Set nextStop = stops.Item(idx)
        If nextStop.ComponentId = currentId Then
            here = idx
        End If
    Next idx
    If here = 0 Then
        idx = IIf(goingForward, 1, stops.Count)
    ElseIf goingForward Then
        idx = here Mod stops.Count + 1
    Else
        idx = (here + stops.Count - 2) Mod stops.Count + 1
    End If
    If idx = here Then Exit Function
    Set nextStop = stops.Item(idx)
    nextStop.TakeFocus Me, False
    nextStop.ScrollIntoSight
    MoveFocusFrom = True
End Function

' Gives keyboard focus to the first control in Tab order.
Public Function FocusFirst() As ReDimUI
    RequireRole ROLE_APP, "FocusFirst"
    MoveFocusFrom vbNullString, True
    Set FocusFirst = Me
End Function

' The button Enter clicks when focus is on a text field or on a control
' that does not use Enter itself. A modal's OK takes over while it is up.
Public Function DefaultButton(ByVal targetComponentId As String) As ReDimUI
    RequireRole ROLE_APP, "DefaultButton"
    mDefaultButtonId = targetComponentId
    Set DefaultButton = Me
End Function

Friend Sub ClickDefaultButton()
    Dim defaultId As String
    Dim defaultTarget As ReDimUI

    RequireRole ROLE_APP, "ClickDefaultButton"
    If ModalUp() Then
        defaultId = "mdl_ok"
    Else
        defaultId = mDefaultButtonId
    End If
    If LenB(defaultId) = 0 Then Exit Sub
    If Not TryGetComponent(defaultId, defaultTarget) Then Exit Sub
    If Not defaultTarget.IsFocusable Then Exit Sub
    ReDimUI.DispatchShape defaultTarget.ShapeName
End Sub

' The access-key letters this app's controls declare.
Friend Function AccessKeyLetters() As String
    Dim componentValue As ReDimUI
    Dim keyLetter As String

    For Each componentValue In mComponents
        keyLetter = componentValue.AccessKeyLetter
        If LenB(keyLetter) > 0 Then
            If InStr(AccessKeyLetters, keyLetter) = 0 Then
                AccessKeyLetters = AccessKeyLetters & keyLetter
            End If
        End If
    Next componentValue
End Function

' Clicks the first control in creation order that declares the key and
' can take a click; while a modal is up, only its buttons count.
Friend Function ActivateAccessKey(ByVal keyLetter As String) As Boolean
    Dim orderedId As Variant
    Dim componentValue As ReDimUI
    Dim trapped As Boolean

    RequireRole ROLE_APP, "ActivateAccessKey"
    trapped = ModalUp()
    For Each orderedId In mComponentOrder
        If TryGetComponent(CStr(orderedId), componentValue) Then
            If componentValue.AccessKeyLetter = keyLetter Then
                If componentValue.IsFocusable Then
                    If Not trapped Or Left$(componentValue.ComponentId, 4) = "mdl_" Then
                        ReDimUI.DispatchShape componentValue.ShapeName
                        ActivateAccessKey = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next orderedId
End Function

' True while Confirm's overlay is showing.
Friend Function ModalUp() As Boolean
    Dim overlay As ReDimUI

    If TryGetComponent("mdl_ov", overlay) Then
        ModalUp = overlay.IsVisible
    End If
End Function

' Esc on a modal: Cancel when it shows one, else OK.
Friend Sub CancelModal()
    Dim modalButton As ReDimUI

    If TryGetComponent("mdl_cancel", modalButton) Then
        If modalButton.IsFocusable Then
            ReDimUI.DispatchShape modalButton.ShapeName
            Exit Sub
        End If
    End If
    If TryGetComponent("mdl_ok", modalButton) Then
        ReDimUI.DispatchShape modalButton.ShapeName
    End If
End Sub

' One open list per app: a list opening, or a click on any other
' control, closes the rest.
Friend Sub CloseListsExcept(ByVal keepComponentId As String)
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "CloseListsExcept"
    For Each componentValue In mComponents
        If componentValue.ComponentId <> keepComponentId Then
            componentValue.CloseDropList
        End If
    Next componentValue
End Sub

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
    If store.Count > 0 Then JoinedSig = "|" & JoinedText(store, "|")
End Function

' A collection's values joined in order: one For Each into an array and
' one Join, where indexed reads and repeated concatenation both went
' quadratic on long lists.
Private Function JoinedText( _
    ByVal store As Collection, _
    ByVal joinWith As String _
) As String
    Dim parts() As String
    Dim entry As Variant
    Dim idx As Long

    If store.Count = 0 Then Exit Function
    ReDim parts(1 To store.Count)
    For Each entry In store
        idx = idx + 1
        parts(idx) = CStr(entry)
    Next entry
    JoinedText = Join(parts, joinWith)
End Function

' Checked item texts in list order, joined for the WritesTo state value.
Private Function CheckedJoined() As String
    Dim parts() As String
    Dim selectedFlags() As Boolean
    Dim entry As Variant
    Dim idx As Long
    Dim takenCount As Long

    If mItems Is Nothing Or mCheckedSet Is Nothing Then Exit Function
    If mItems.Count = 0 Or mCheckedSet.Count = 0 Then Exit Function
    ReDim selectedFlags(1 To mItems.Count)
    For Each entry In mCheckedSet
        If CLng(entry) >= 1 And CLng(entry) <= mItems.Count Then
            selectedFlags(CLng(entry)) = True
        End If
    Next entry
    EnsureItemTexts
    ReDim parts(1 To mItems.Count)
    For idx = 1 To mItems.Count
        If selectedFlags(idx) Then
            takenCount = takenCount + 1
            parts(takenCount) = mItemTexts(idx)
        End If
    Next idx
    If takenCount = 0 Then Exit Function
    ReDim Preserve parts(1 To takenCount)
    CheckedJoined = Join(parts, ", ")
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
    Dim kept As Collection
    Dim selectedFlags() As Boolean
    Dim entry As Variant
    Dim idx As Long
    Dim anySelected As Boolean

    If rightward Then
        Set sourceItems = mItems
        Set targetItems = mChosenItems
        Set selectionSet = mSelLeft
    Else
        Set sourceItems = mChosenItems
        Set targetItems = mItems
        Set selectionSet = mSelRight
    End If
    If sourceItems.Count = 0 Or selectionSet.Count = 0 Then Exit Sub
    ReDim selectedFlags(1 To sourceItems.Count)
    For Each entry In selectionSet
        If CLng(entry) >= 1 And CLng(entry) <= sourceItems.Count Then
            selectedFlags(CLng(entry)) = True
            anySelected = True
        End If
    Next entry
    If Not anySelected Then Exit Sub
    ' One pass moves the selected rows to the target in their list order
    ' and rebuilds the source without them; removing by index would walk
    ' the list once per row.
    Set kept = New Collection
    For Each entry In sourceItems
        idx = idx + 1
        If selectedFlags(idx) Then
            targetItems.Add CStr(entry)
        Else
            kept.Add CStr(entry)
        End If
    Next entry
    If rightward Then
        Set mItems = kept
        Set mSelLeft = New Collection
    Else
        Set mChosenItems = kept
        Set mSelRight = New Collection
    End If
    ItemsChanged
    Touch
    CommitTransferChange hostApp
End Sub

Private Sub TransferAll(ByVal hostApp As ReDimUI, ByVal rightward As Boolean)
    Dim entry As Variant

    If rightward Then
        If mItems.Count = 0 Then Exit Sub
        For Each entry In mItems
            mChosenItems.Add CStr(entry)
        Next entry
        Set mItems = New Collection
    Else
        If mChosenItems.Count = 0 Then Exit Sub
        For Each entry In mChosenItems
            mItems.Add CStr(entry)
        Next entry
        Set mChosenItems = New Collection
    End If
    ItemsChanged
    Set mSelLeft = New Collection
    Set mSelRight = New Collection
    Touch
    CommitTransferChange hostApp
End Sub

' A user transfer writes the chosen list (joined with a comma and space)
' to state and fires OnChange once. Selection clicks fire nothing.
Private Sub CommitTransferChange(ByVal hostApp As ReDimUI)
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, JoinedText(mChosenItems, ", ")
    End If
    If LenB(mClickProc) > 0 Then
        On Error Resume Next
        Application.Run mClickProc
        On Error GoTo 0
    End If
End Sub

Friend Function CanClick() As Boolean
    If Not mVisible Or Not mEnabled Or mBusy Then Exit Function
    ' A float field takes every click, so the second of a double click
    ' reaches it; a click there only moves the caret.
    If Not IsFloatField() Then
        If GetTickCount64() - mLastClickMs < CLICK_DEBOUNCE_MS Then Exit Function
    End If
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
            If ReDimUI.SenderPart = "cx" Then
                ClearFieldText hostApp
            ElseIf IsFloatField() Then
                ClickField hostApp
            Else
                ' The frame covers the anchor cell and takes the click, so
                ' the click selects that cell: typing lands in the grid and
                ' F2 edits in place, as if the cell itself was clicked.
                SelectInputCell hostApp.Sheet
            End If
            Exit Sub
        Case KIND_COMBO
            If ReDimUI.SenderPart = "cx" Then
                ClearFieldText hostApp
                Exit Sub
            ElseIf ReDimUI.SenderPart = "optn" Then
                ' The empty row is a message, not an option.
                Exit Sub
            ElseIf ReDimUI.SenderPart = "optu" Or ReDimUI.SenderPart = "optd" Then
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
                ' The caret glyph toggles the list; the face takes the
                ' caret like any field.
                If ReDimUI.SenderPart = "caret" _
                    And ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
                    ToggleDropList hostApp.Sheet
                Else
                    ClickField hostApp
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
            If ReDimUI.SenderPart = "optn" Then
                Exit Sub
            ElseIf ReDimUI.SenderPart = "optu" Or ReDimUI.SenderPart = "optd" Then
                PageDropList (ReDimUI.SenderPart = "optd")
                Exit Sub
            ElseIf Left$(ReDimUI.SenderPart, 3) = "opt" _
                And IsNumeric(Mid$(ReDimUI.SenderPart, 4)) Then
                Dim pickedItem As Long
                pickedItem = ListItemAtRow(CLng(Mid$(ReDimUI.SenderPart, 4)))
                If pickedItem = 0 Then Exit Sub
                ' Re-picking the selection only closes the list: nothing
                ' changed, so nothing is written and no handler runs.
                If Not PickSelectItem(hostApp, pickedItem) Then Exit Sub
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
            If checkPart = "mb" Or checkPart = "mt" Then
                checkRow = 0
            ElseIf LenB(checkPart) = 0 Then
                checkRow = 1
            ElseIf IsNumeric(Mid$(checkPart, 2)) Then
                checkRow = CLng(Mid$(checkPart, 2))
            Else
                Exit Sub
            End If
            If Not ToggleCheckRow(hostApp, checkRow) Then Exit Sub
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
                    PlaceKeyCursor 0, transferRow
                    Touch
                End If
            ElseIf Left$(transferPart, 2) = "cl" Then
                If IsNumeric(Mid$(transferPart, 3)) Then
                    transferRow = mRightOffset + CLng(Mid$(transferPart, 3))
                    TransferSelectionToggle mSelRight, transferRow
                    PlaceKeyCursor 1, transferRow
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
            If Not SelectRadioRow(hostApp, rowNumber) Then Exit Sub
        Case KIND_STEPPER
            Dim stepped As Double
            If ReDimUI.SenderPart = "minus" Then
                stepped = mValue - mSliderStep
            ElseIf ReDimUI.SenderPart = "plus" Then
                stepped = mValue + mSliderStep
            Else
                Exit Sub
            End If
            If Not SetStepperValue(hostApp, stepped) Then Exit Sub
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
            ' Any click dismisses; the action button also runs its
            ' handler, whose errors reach the app's error handler.
            mExpiresAtMs = 1
            ReDimUI.RequestPump
            If ReDimUI.SenderPart = "ta" And LenB(mToastActionProc) > 0 Then
                Application.Run mToastActionProc
            End If
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
    EnsureItemTexts
    ItemAt = mItemTexts(itemNumber)
End Function

' Called after every change to mItems: the text snapshot and the cached
' list matches are rebuilt on next use.
Private Sub ItemsChanged()
    mItemTextsCurrent = False
    Set mListMatches = Nothing
End Sub

Private Sub EnsureItemTexts()
    Dim entry As Variant
    Dim idx As Long

    If mItemTextsCurrent Then Exit Sub
    Erase mItemTexts
    If Not mItems Is Nothing Then
        If mItems.Count > 0 Then
            ReDim mItemTexts(1 To mItems.Count)
            For Each entry In mItems
                idx = idx + 1
                mItemTexts(idx) = CStr(entry)
            Next entry
        End If
    End If
    mItemTextsCurrent = True
End Sub

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
    Dim wasApplied As Boolean
    Dim growBefore As Double

    RequireRole ROLE_COMPONENT, "Apply"
    ResolveBindings hostApp
    wasApplied = mAppliedValid
    growBefore = mGrowAdded
    EnsureGeometryResolved hostSheet
    Set mainShape = EnsureShape(hostSheet)
    ApplyGeometry mainShape
    ApplyStyle mainShape, themeValue
    ApplyText mainShape, themeValue
    ApplyBehavior mainShape, hostSheet
    ApplyKindExtras mainShape, hostSheet, themeValue
    If IsFloatField() Then
        ApplyClearButton mainShape, hostSheet, themeValue
        ApplyFieldNotes hostSheet, themeValue
    End If
    ApplyFocusRing hostSheet, themeValue
    ApplyAltText mainShape
    mAppliedValid = True
    ' A field that grew or shrank since it was drawn moves what is placed
    ' against it; a first draw leaves that to the render in progress.
    If wasApplied And Abs(mGrowAdded - growBefore) > 0.01 Then
        hostApp.ReapplyDependents mComponentId
    End If
End Sub

Private Sub ApplyAltText(ByVal mainShape As Shape)
    Dim described As String

    described = AccessibleDescription()
    If described = mAppliedAltText Then
        If mAppliedValid Or LenB(described) = 0 Then Exit Sub
    End If
    mAppliedAltText = described
    On Error Resume Next
    mainShape.AlternativeText = described
    On Error GoTo 0
End Sub

' What a screen reader says for the control: the AltText override, or
' its text, its role, and the state that matters for the kind. Labels,
' cards, and toasts are their text, which a reader already gets, so
' they carry none; typed field text stays out for the same reason, and
' progress reads in 5 percent steps. Each of those would otherwise
' cost a write on the hottest update paths.
Private Function AccessibleDescription() As String
    Dim described As String

    If LenB(mAltText) > 0 Then
        AccessibleDescription = mAltText
        Exit Function
    End If
    Select Case mKind
        Case KIND_BUTTON
            described = mText & ", button"
            If mBusy Then described = described & ", busy"
        Case KIND_PROGRESS
            described = "Progress, " & CStr(CLng(mValue / 5) * 5) & " percent"
        Case KIND_SPINNER
            described = "Busy"
        Case KIND_TOGGLE
            described = "Switch, " & IIf(mChecked, "on", "off")
        Case KIND_INPUT
            described = "Edit field"
            If LenB(mPlaceholder) > 0 Then described = described & ", " & mPlaceholder
        Case KIND_COMBO
            described = "Combo box"
            If LenB(mPlaceholder) > 0 Then described = described & ", " & mPlaceholder
        Case KIND_SELECT
            If CLng(mValue) >= 1 And CLng(mValue) <= ItemCount Then
                described = "Drop-down, " & ItemAt(CLng(mValue))
            Else
                described = "Drop-down, " & mText
            End If
        Case KIND_TICKBOX
            described = mText & ", checkbox, " & _
                IIf(mChecked, "checked", "not checked")
        Case KIND_RADIO
            described = "Radio group"
            If CLng(mValue) >= 1 And CLng(mValue) <= ItemCount Then
                described = described & ", " & ItemAt(CLng(mValue)) & " selected"
            End If
        Case KIND_STEPPER
            described = "Stepper, " & CStr(mValue)
        Case KIND_SLIDEBAR
            described = "Slider, " & CStr(mValue)
        Case KIND_TRANSFER
            described = "Transfer list"
        Case KIND_CHECKLIST
            EnsureCheckedSet
            described = "Checklist, " & CStr(mCheckedSet.Count) & " of " & _
                CStr(ItemCount) & " checked"
        Case KIND_IMAGE
            described = "Image"
    End Select
    If Not mEnabled And LenB(described) > 0 Then
        described = described & ", unavailable"
    End If
    If LenB(mAccessKey) > 0 And LenB(described) > 0 Then
        described = described & ", Alt+" & UCase$(mAccessKey)
    End If
    AccessibleDescription = described
End Function

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
    mGrowAdded = 0
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
    If mAutoGrowLines > 0 Or mGrowAdded <> 0 Then GrowToText
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

' The component this one is placed Below or RightOf, or "" for none.
Friend Property Get PlacedRelativeTo() As String
    If mRelMode <> REL_NONE Then PlacedRelativeTo = mRelTargetId
End Property

' AutoGrow: a multi-line float field's height follows the lines its text
' takes, never below the height it was given and never past the cap.
Private Sub GrowToText()
    Dim resolvedSize As Double
    Dim baseHeight As Double
    Dim neededHeight As Double
    Dim lineTotal As Long

    baseHeight = mHeight - mGrowAdded
    neededHeight = baseHeight
    If mAutoGrowLines > 0 And mMultiLine And IsFloatField() Then
        resolvedSize = mFontSize
        If resolvedSize <= 0 Then resolvedSize = OwnerApp.Theme.BaseFontSize
        lineTotal = WrappedLineCount(FaceCharsPerLine(resolvedSize))
        If lineTotal > mAutoGrowLines Then lineTotal = mAutoGrowLines
        If lineTotal * FaceLineHeight(resolvedSize) + 6 > neededHeight Then
            neededHeight = lineTotal * FaceLineHeight(resolvedSize) + 6
        End If
    End If
    mGrowAdded = neededHeight - baseHeight
    mHeight = neededHeight
End Sub

' The lines the text takes at rest, where a line longer than the face
' holds wraps. The height counts them with focus too, though the editing
' face scrolls a long line instead, so focus never resizes the field.
Private Function WrappedLineCount(ByVal charsPerLine As Long) As Long
    Dim textLines() As String
    Dim lineNo As Long
    Dim lineLength As Long

    If LenB(mText) = 0 Then
        WrappedLineCount = 1
        Exit Function
    End If
    textLines = Split(mText, vbLf)
    For lineNo = LBound(textLines) To UBound(textLines)
        lineLength = Len(textLines(lineNo))
        If lineLength <= charsPerLine Then
            WrappedLineCount = WrappedLineCount + 1
        Else
            WrappedLineCount = WrappedLineCount + _
                (lineLength + charsPerLine - 1) \ charsPerLine
        End If
    Next lineNo
End Function

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
        ForgetPartSignatures
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

' The named part, created when missing; wasCreated says which, so a row
' updater that writes only what changed knows to style a fresh shape.
Private Function EnsurePartShape( _
    ByVal hostSheet As Worksheet, _
    ByVal partName As String, _
    ByVal shapeType As MsoAutoShapeType, _
    Optional ByRef wasCreated As Boolean _
) As Shape
    Dim partShape As Shape
    Dim partShapeName As String

    partShapeName = ShapeName & PART_SEPARATOR & partName
    On Error Resume Next
    Set partShape = hostSheet.Shapes(partShapeName)
    On Error GoTo 0
    wasCreated = (partShape Is Nothing)
    If wasCreated Then
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
                Dim errorOn As Boolean
                ringOn = ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId)
                errorOn = LenB(mValidationError) > 0
                If Not mAppliedValid _
                    Or mAppliedFill <> ResolveSurfaceFill(themeValue) _
                    Or mAppliedFocusRing <> ringOn _
                    Or mAppliedErrorBorder <> errorOn Then
                    mainShape.Fill.Visible = msoTrue
                    mainShape.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
                    mainShape.Line.Visible = msoTrue
                    mAppliedErrorBorder = errorOn
                    If errorOn Then
                        mainShape.Line.ForeColor.RGB = themeValue.DangerColor
                        mainShape.Line.Weight = 1.5
                    ElseIf ringOn Then
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
        Case VARIANT_WARNING
            ResolveFill = themeValue.WarningColor
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
    If mKind = KIND_TOAST Then
        If LenB(ToastIcon()) > 0 Then
            ResolveText = ToastIcon() & vbTab & mText
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
' The focused field's face: the text with the insertion bar at the
' caret (a space while the blink is off, so nothing shifts), windowed
' to what fits, with ellipses marking what is cut: the lines around the
' caret, and a long caret line trimmed around the caret. Fills mFaceMap,
' which selection painting and click-to-caret read.
Private Function BuildFace( _
    ByVal resolvedSize As Double, _
    Optional ByVal ghostText As String _
) As String
    Dim rendered As String
    Dim renderMap() As Long
    Dim caretRender As Long
    Dim idx As Long
    Dim charsPerLine As Long
    Dim maxLines As Long
    Dim lineStarts() As Long
    Dim lineEnds() As Long
    Dim lineTotal As Long
    Dim segmentStart As Long
    Dim caretLine As Long
    Dim firstLine As Long
    Dim lastLine As Long
    Dim lineNumber As Long
    Dim lineLength As Long
    Dim windowStart As Long
    Dim faceText As String

    If mCaretPos > Len(mText) Then mCaretPos = Len(mText)
    If mSelAnchor < 0 Or mSelAnchor > Len(mText) Then mSelAnchor = mCaretPos
    caretRender = mCaretPos + 1
    ' Completion letters follow the text and map to -1.
    rendered = Left$(mText, mCaretPos) & IIf(mCaretOn, "|", " ") & _
        Mid$(mText, mCaretPos + 1) & ghostText
    ReDim renderMap(1 To Len(rendered))
    For idx = 1 To Len(rendered)
        If idx < caretRender Then
            renderMap(idx) = idx
        ElseIf idx > Len(mText) + 1 Then
            renderMap(idx) = -1
        ElseIf idx > caretRender Then
            renderMap(idx) = idx - 1
        End If
    Next idx
    charsPerLine = FaceCharsPerLine(resolvedSize)
    ' Line spans of the rendered text, line breaks excluded.
    segmentStart = 1
    For idx = 1 To Len(rendered) + 1
        If idx > Len(rendered) Or Mid$(rendered, idx, 1) = vbLf Then
            lineTotal = lineTotal + 1
            ReDim Preserve lineStarts(1 To lineTotal)
            ReDim Preserve lineEnds(1 To lineTotal)
            lineStarts(lineTotal) = segmentStart
            lineEnds(lineTotal) = idx - 1
            If caretRender >= segmentStart And caretRender <= idx - 1 Then
                caretLine = lineTotal
            End If
            segmentStart = idx + 1
        End If
    Next idx
    If caretLine = 0 Then caretLine = lineTotal
    firstLine = 1
    lastLine = lineTotal
    If mMultiLine Then
        maxLines = Int((mHeight - 6) / FaceLineHeight(resolvedSize) + 0.001)
        If maxLines < 1 Then maxLines = 1
        firstLine = caretLine - maxLines + 1
        If firstLine < 1 Then firstLine = 1
        lastLine = firstLine + maxLines - 1
        If lastLine > lineTotal Then lastLine = lineTotal
    End If
    mFaceLength = 0
    mGhostFaceStart = 0
    mGhostFaceLength = 0
    ReDim mFaceMap(1 To Len(rendered) + 2 * lineTotal + 2)
    If firstLine > 1 Then AppendFace faceText, ChrW(8230), 0
    For lineNumber = firstLine To lastLine
        If lineNumber > firstLine Then
            AppendFace faceText, vbLf, renderMap(lineStarts(lineNumber) - 1)
        End If
        lineLength = lineEnds(lineNumber) - lineStarts(lineNumber) + 1
        If lineLength <= charsPerLine Then
            For idx = lineStarts(lineNumber) To lineEnds(lineNumber)
                AppendFace faceText, Mid$(rendered, idx, 1), renderMap(idx)
            Next idx
        ElseIf lineNumber = caretLine Then
            windowStart = caretRender - lineStarts(lineNumber) + 1 - charsPerLine + 2
            If windowStart < 1 Then windowStart = 1
            If windowStart + charsPerLine - 1 > lineLength Then
                windowStart = lineLength - charsPerLine + 1
            End If
            If windowStart > 1 Then AppendFace faceText, ChrW(8230), 0
            For idx = lineStarts(lineNumber) + windowStart - 1 To _
                lineStarts(lineNumber) + windowStart + charsPerLine - 2
                AppendFace faceText, Mid$(rendered, idx, 1), renderMap(idx)
            Next idx
            If windowStart + charsPerLine - 1 < lineLength Then
                AppendFace faceText, ChrW(8230), 0
            End If
        Else
            For idx = lineStarts(lineNumber) To lineStarts(lineNumber) + charsPerLine - 2
                AppendFace faceText, Mid$(rendered, idx, 1), renderMap(idx)
            Next idx
            AppendFace faceText, ChrW(8230), 0
        End If
    Next lineNumber
    If lastLine < lineTotal Then AppendFace faceText, ChrW(8230), 0
    ' An empty field keeps its placeholder, after the bar.
    If LenB(mText) = 0 And LenB(mPlaceholder) > 0 Then
        For idx = 1 To Len(mPlaceholder)
            AppendFace faceText, Mid$(mPlaceholder, idx, 1), 0
        Next idx
    End If
    BuildFace = faceText
End Function

Private Sub AppendFace(ByRef faceText As String, ByVal faceChar As String, ByVal textIndex As Long)
    faceText = faceText & faceChar
    mFaceLength = mFaceLength + 1
    If mFaceLength > UBound(mFaceMap) Then
        ReDim Preserve mFaceMap(1 To mFaceLength + 8)
    End If
    ' Completion letters show no text character; the face notes where
    ' they run, for their muted ink.
    If textIndex < 0 Then
        If mGhostFaceLength = 0 Then mGhostFaceStart = mFaceLength
        mGhostFaceLength = mGhostFaceLength + 1
        textIndex = 0
    End If
    mFaceMap(mFaceLength) = textIndex
End Sub

' The characters a face line holds and the height a line takes, both
' estimated from the font size because they are needed before the text
' is written.
Private Function FaceCharsPerLine(ByVal resolvedSize As Double) As Long
    FaceCharsPerLine = Int((mWidth - 12 - ClearReserve()) / (resolvedSize * 0.58))
    If FaceCharsPerLine < 4 Then FaceCharsPerLine = 4
End Function

Private Function FaceLineHeight(ByVal resolvedSize As Double) As Double
    FaceLineHeight = resolvedSize * 1.35
End Function

' The selection shows as a highlight in the accent with accent ink. Excel
' offers no way to take a highlight off, and a text write spreads the
' first character's formatting over the rest, so once a field has shown
' a selection its face is reset (the fill as highlight, the resolved
' ink) whenever marks change or the face is rewritten or restyled.
' Completion letters read in muted ink, painted again after every
' rewrite for the same reason.
Private Sub PaintFieldMarks( _
    ByVal mainShape As Shape, _
    ByVal themeValue As ReDimUI, _
    ByVal inkColor As Long, _
    ByVal focusedField As Boolean, _
    ByVal faceRewritten As Boolean _
)
    Dim marksWanted As Boolean
    Dim resetWanted As Boolean
    Dim ghostKey As String
    Dim selFrom As Long
    Dim selTo As Long
    Dim idx As Long
    Dim runStart As Long
    Dim inRun As Boolean

    If mSelAnchor < 0 Or mSelAnchor > Len(mText) Then mSelAnchor = mCaretPos
    marksWanted = focusedField And HasSelection()
    If focusedField And mGhostFaceLength > 0 Then
        ghostKey = CStr(mGhostFaceStart) & "+" & CStr(mGhostFaceLength)
    End If
    resetWanted = marksWanted Or mMarksPainted _
        Or (mHighlightTouched And faceRewritten)
    If Not resetWanted And ghostKey = mAppliedGhostKey Then
        If Not faceRewritten Or LenB(ghostKey) = 0 Then Exit Sub
    End If
    On Error Resume Next
    With mainShape.TextFrame2.TextRange
        If resetWanted Then
            .Font.Highlight.RGB = ResolveSurfaceFill(themeValue)
            .Font.Fill.ForeColor.RGB = inkColor
            If marksWanted Then
                selFrom = SelStart() + 1
                selTo = SelStart() + SelLength()
                For idx = 1 To mFaceLength + 1
                    inRun = False
                    If idx <= mFaceLength Then
                        inRun = (mFaceMap(idx) >= selFrom And mFaceMap(idx) <= selTo)
                    End If
                    If inRun Then
                        If runStart = 0 Then runStart = idx
                    ElseIf runStart > 0 Then
                        With .Characters(runStart, idx - runStart).Font
                            .Highlight.RGB = themeValue.PrimaryColor
                            .Fill.ForeColor.RGB = themeValue.OnPrimaryColor
                        End With
                        runStart = 0
                    End If
                Next idx
            End If
        ElseIf LenB(mAppliedGhostKey) > 0 And Not faceRewritten Then
            ' Completion letters that went without a rewrite give the
            ' ink back.
            .Font.Fill.ForeColor.RGB = inkColor
        End If
        If LenB(ghostKey) > 0 Then
            .Characters(mGhostFaceStart, mGhostFaceLength).Font.Fill _
                .ForeColor.RGB = themeValue.OnMutedColor
        End If
    End With
    On Error GoTo 0
    If resetWanted Then
        mHighlightTouched = True
        mMarksPainted = marksWanted
    End If
    mAppliedGhostKey = ghostKey
End Sub

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
    ' Shape text cannot scroll, so a focused field renders a window of
    ' its buffer that follows the caret. The full text stays in the
    ' buffer and commits intact.
    Dim focusedField As Boolean
    If IsFloatField() Then
        focusedField = ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId)
        If focusedField Then
            If mKind = KIND_COMBO Then
                shownText = BuildFace(resolvedSize, GhostRest(mainShape.Parent))
            Else
                shownText = BuildFace(resolvedSize)
            End If
        ElseIf LenB(mText) = 0 And LenB(mPlaceholder) > 0 Then
            shownText = mPlaceholder
        End If
        ' A placeholder reads in muted ink.
        If LenB(mText) = 0 And LenB(mPlaceholder) > 0 Then
            resolvedColor = themeValue.OnMutedColor
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
    If IsFloatField() Then
        PaintFieldMarks mainShape, themeValue, resolvedColor, focusedField, _
            textChanged Or mustRestyle
    End If
    If mKind = KIND_BUTTON Then
        If LenB(mAccessKey) > 0 Or LenB(mAppliedAccessKey) > 0 Then
            If textChanged Or mustRestyle Or mAppliedAccessKey <> mAccessKey Then
                MarkAccessKey mainShape, shownText
            End If
        End If
    End If
End Sub

' Underlines the access key's first appearance in the shown text.
Private Sub MarkAccessKey(ByVal keyedShape As Shape, ByVal shownText As String)
    Dim keyAt As Long

    mAppliedAccessKey = mAccessKey
    On Error Resume Next
    With keyedShape.TextFrame2.TextRange
        .Font.UnderlineStyle = msoNoUnderline
        If LenB(mAccessKey) > 0 Then
            keyAt = InStr(1, shownText, mAccessKey, vbTextCompare)
            If keyAt > 0 Then
                .Characters(keyAt, 1).Font.UnderlineStyle = msoUnderlineSingleLine
            End If
        End If
    End With
    On Error GoTo 0
End Sub

' Keys for a focused control that is not a text field. Tab and Shift+Tab
' move focus; Esc closes an open list, cancels a modal, or leaves; the
' rest is per kind, with Enter clicking a button or the app's default
' button.
Private Sub HandleControlKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Select Case keyText
        Case "{TAB}", "{BACKTAB}"
            If mKind = KIND_SELECT And mSelectOpen Then
                ' Tab takes the highlighted item, as a native list does.
                If mListHighlight >= 1 Then
                    If PickSelectItem(hostApp, mListHighlight) Then
                        RunClickHandlers hostApp
                    End If
                Else
                    CloseDropList
                End If
            End If
            If Not hostApp.MoveFocusFrom(mComponentId, keyText = "{TAB}") Then
                LeaveFocus hostApp, True
            End If
            Exit Sub
        Case "{ESC}"
            If mSelectOpen Then
                CloseDropList
            ElseIf hostApp.ModalUp Then
                hostApp.CancelModal
            Else
                LeaveFocus hostApp, True
            End If
            Exit Sub
    End Select
    Select Case mKind
        Case KIND_BUTTON, KIND_IMAGE
            If keyText = " " Or keyText = "{ENTER}" Then ActivateByKey
        Case KIND_TOGGLE, KIND_TICKBOX
            If keyText = " " Then
                ActivateByKey
            ElseIf keyText = "{ENTER}" Then
                hostApp.ClickDefaultButton
            End If
        Case KIND_RADIO
            HandleRadioKey hostApp, keyText
        Case KIND_STEPPER
            HandleStepperKey hostApp, keyText
        Case KIND_SLIDEBAR
            HandleSliderKey hostApp, keyText
        Case KIND_SELECT
            HandleSelectKey hostApp, keyText
        Case KIND_CHECKLIST
            HandleCheckListKey hostApp, keyText
        Case KIND_TRANSFER
            HandleTransferKey hostApp, keyText
    End Select
End Sub

' A key standing in for a click takes the mouse's own path: debounce,
' handlers, and the sender context included.
Private Sub ActivateByKey()
    ReDimUI.DispatchShape ShapeName
End Sub

' Arrows move the selection itself, wrapping at the ends; each move
' writes the state and fires OnChange as a click would.
Private Sub HandleRadioKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim itemTotal As Long
    Dim aimRow As Long

    itemTotal = ItemCount
    If itemTotal = 0 Then Exit Sub
    Select Case keyText
        Case "{UP}", "{LEFT}"
            aimRow = CLng(mValue) - 1
            If aimRow < 1 Then aimRow = itemTotal
        Case "{DOWN}", "{RIGHT}"
            aimRow = CLng(mValue) + 1
            If aimRow > itemTotal Then aimRow = 1
        Case "{HOME}"
            aimRow = 1
        Case "{END}"
            aimRow = itemTotal
        Case " "
            If CLng(mValue) = 0 Then aimRow = 1
        Case "{ENTER}"
            hostApp.ClickDefaultButton
    End Select
    If aimRow = 0 Then Exit Sub
    If SelectRadioRow(hostApp, aimRow) Then RunClickHandlers hostApp
End Sub

Private Function SelectRadioRow(ByVal hostApp As ReDimUI, ByVal rowNumber As Long) As Boolean
    If rowNumber < 1 Or rowNumber > ItemCount Then Exit Function
    If rowNumber = CLng(mValue) Then Exit Function
    mValue = rowNumber
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, ItemAt(rowNumber)
    End If
    Touch
    SelectRadioRow = True
End Function

' Arrows step, Page Up and Page Down step ten times, Home and End jump
' to the range ends; each change fires OnChange.
Private Sub HandleStepperKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim aimValue As Double

    Select Case keyText
        Case "{UP}", "{RIGHT}"
            aimValue = mValue + mSliderStep
        Case "{DOWN}", "{LEFT}"
            aimValue = mValue - mSliderStep
        Case "{PGUP}"
            aimValue = mValue + 10 * mSliderStep
        Case "{PGDN}"
            aimValue = mValue - 10 * mSliderStep
        Case "{HOME}"
            aimValue = mSliderMin
        Case "{END}"
            aimValue = mSliderMax
        Case "{ENTER}"
            hostApp.ClickDefaultButton
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If SetStepperValue(hostApp, aimValue) Then RunClickHandlers hostApp
End Sub

' Clamps to the range and, when the value moves, writes it and its
' state. True when it moved; the caller runs the handlers.
Private Function SetStepperValue(ByVal hostApp As ReDimUI, ByVal aimValue As Double) As Boolean
    If aimValue < mSliderMin Then aimValue = mSliderMin
    If aimValue > mSliderMax Then aimValue = mSliderMax
    If aimValue = mValue Then Exit Function
    mValue = aimValue
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mValue
    Touch
    SetStepperValue = True
End Function

' Arrows move a step, Page Up and Page Down a tenth of the range (in
' whole steps), Home and End the ends; each change fires OnChange.
Private Sub HandleSliderKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim aimValue As Double
    Dim smallStep As Double
    Dim pageStep As Double

    If mSliderMax <= mSliderMin Then Exit Sub
    smallStep = mSliderStep
    If smallStep <= 0 Then smallStep = (mSliderMax - mSliderMin) / 100
    pageStep = CLng((mSliderMax - mSliderMin) / 10 / smallStep) * smallStep
    If pageStep < smallStep Then pageStep = smallStep
    Select Case keyText
        Case "{RIGHT}", "{UP}"
            aimValue = mValue + smallStep
        Case "{LEFT}", "{DOWN}"
            aimValue = mValue - smallStep
        Case "{PGUP}"
            aimValue = mValue + pageStep
        Case "{PGDN}"
            aimValue = mValue - pageStep
        Case "{HOME}"
            aimValue = mSliderMin
        Case "{END}"
            aimValue = mSliderMax
        Case "{ENTER}"
            hostApp.ClickDefaultButton
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If aimValue < mSliderMin Then aimValue = mSliderMin
    If aimValue > mSliderMax Then aimValue = mSliderMax
    If aimValue = mValue Then Exit Sub
    SlideToFraction hostApp, (aimValue - mSliderMin) / (mSliderMax - mSliderMin), True
End Sub

' A focused select. Closed, arrows and Home, End, Page Up, and Page
' Down change the selection directly, as a native drop-down does, and
' Space, Alt+Down, and F4 open the list. Open, they move the highlight,
' Enter, Space, or Alt+Up take it, and F4 closes. Typed letters jump to
' the next item starting with them, open or closed.
Private Sub HandleSelectKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim itemTotal As Long
    Dim aimRow As Long
    Dim anchorItem As Long

    itemTotal = ItemCount
    If itemTotal = 0 Then
        If keyText = "{ENTER}" Then hostApp.ClickDefaultButton
        Exit Sub
    End If
    If mSelectOpen Then
        anchorItem = mListHighlight
        If anchorItem < 1 Then anchorItem = CLng(mValue)
        Select Case keyText
            Case "{UP}"
                aimRow = anchorItem - 1
            Case "{DOWN}"
                aimRow = anchorItem + 1
            Case "{PGUP}"
                aimRow = anchorItem - mListRows
            Case "{PGDN}"
                aimRow = anchorItem + mListRows
            Case "{HOME}"
                aimRow = 1
            Case "{END}"
                aimRow = itemTotal
            Case "{ENTER}", " ", "{ALTUP}"
                If mListHighlight >= 1 Then
                    If PickSelectItem(hostApp, mListHighlight) Then
                        RunClickHandlers hostApp
                    End If
                Else
                    CloseDropList
                End If
                Exit Sub
            Case "{F4}", "{ALTDOWN}"
                CloseDropList
                Exit Sub
            Case Else
                aimRow = TypeAheadItem(keyText, anchorItem)
                If aimRow = 0 Then Exit Sub
        End Select
        If aimRow < 1 Then aimRow = 1
        If aimRow > itemTotal Then aimRow = itemTotal
        HighlightListItem aimRow
        Exit Sub
    End If
    anchorItem = CLng(mValue)
    Select Case keyText
        Case "{UP}", "{LEFT}"
            aimRow = anchorItem - 1
        Case "{DOWN}", "{RIGHT}"
            aimRow = anchorItem + 1
        Case "{PGUP}"
            aimRow = anchorItem - mListRows
        Case "{PGDN}"
            aimRow = anchorItem + mListRows
        Case "{HOME}"
            aimRow = 1
        Case "{END}"
            aimRow = itemTotal
        Case " ", "{ALTDOWN}", "{F4}"
            OpenDropList hostApp.Sheet
            HighlightListItem IIf(anchorItem >= 1, anchorItem, 1)
            Exit Sub
        Case "{ENTER}"
            hostApp.ClickDefaultButton
            Exit Sub
        Case Else
            aimRow = TypeAheadItem(keyText, anchorItem)
            If aimRow = 0 Then Exit Sub
    End Select
    If aimRow < 1 Then aimRow = 1
    If aimRow > itemTotal Then aimRow = itemTotal
    If PickSelectItem(hostApp, aimRow) Then RunClickHandlers hostApp
End Sub

' Moves an open select's highlight, scrolling the window to keep it in
' view.
Private Sub HighlightListItem(ByVal itemIndex As Long)
    mListHighlight = itemIndex
    If mListHighlight > mListOffset + mListRows Then
        mListOffset = mListHighlight - mListRows
    ElseIf mListHighlight <= mListOffset Then
        mListOffset = mListHighlight - 1
    End If
    Touch
End Sub

' Takes an item into the select and closes its list. True when the
' value changed, which is when the caller runs the handlers; taking the
' current item again only closes the list.
Private Function PickSelectItem(ByVal hostApp As ReDimUI, ByVal itemIndex As Long) As Boolean
    mSelectOpen = False
    mListHighlight = 0
    If itemIndex = CLng(mValue) Then
        Touch
        Exit Function
    End If
    mValue = itemIndex
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, ItemAt(itemIndex)
    End If
    Touch
    PickSelectItem = True
End Function

' Type-ahead: letters typed within a second of each other build a
' prefix, and the answer is the next item that starts with it, counted
' from the current one; one letter typed again steps through the items
' that start with it. 0 when nothing matches.
Private Function TypeAheadItem(ByVal keyText As String, ByVal fromItem As Long) As Long
    Dim nowMs As LongLong
    Dim typedPrefix As String
    Dim itemTotal As Long
    Dim startAt As Long
    Dim stepCount As Long
    Dim candidate As Long

    If Len(keyText) <> 1 Then Exit Function
    nowMs = GetTickCount64()
    If nowMs - mTypeAheadMs > 1000 Then mTypeAhead = vbNullString
    mTypeAheadMs = nowMs
    mTypeAhead = mTypeAhead & LCase$(keyText)
    typedPrefix = mTypeAhead
    If Len(typedPrefix) > 1 And typedPrefix = String$(Len(typedPrefix), Left$(typedPrefix, 1)) Then
        typedPrefix = Left$(typedPrefix, 1)
    End If
    itemTotal = ItemCount
    ' One letter looks past the current item; a longer prefix refines
    ' from it, so typing on narrows instead of skipping ahead.
    If Len(typedPrefix) = 1 Then
        startAt = fromItem + 1
    Else
        startAt = fromItem
    End If
    If startAt < 1 Then startAt = 1
    For stepCount = 0 To itemTotal - 1
        candidate = (startAt - 1 + stepCount) Mod itemTotal + 1
        If LCase$(Left$(ItemAt(candidate), Len(typedPrefix))) = typedPrefix Then
            TypeAheadItem = candidate
            Exit Function
        End If
    Next stepCount
End Function

' Up and Down move a cursor through the rows, the select-all header
' included; Space toggles the cursor's row as a click would.
Private Sub HandleCheckListKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim firstRow As Long
    Dim aimRow As Long

    If ItemCount = 0 Then Exit Sub
    firstRow = IIf(mSelectAllRow, 0, 1)
    Select Case keyText
        Case "{UP}"
            aimRow = mKeyRow - 1
        Case "{DOWN}"
            aimRow = mKeyRow + 1
        Case "{HOME}", "{PGUP}"
            aimRow = firstRow
        Case "{END}", "{PGDN}"
            aimRow = ItemCount
        Case " "
            If ToggleCheckRow(hostApp, mKeyRow) Then RunClickHandlers hostApp
            Exit Sub
        Case "{ENTER}"
            hostApp.ClickDefaultButton
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If aimRow < firstRow Then aimRow = firstRow
    If aimRow > ItemCount Then aimRow = ItemCount
    PlaceKeyCursor 0, aimRow
    Touch
End Sub

' Toggles one row, or with row 0 the select-all header: from mixed or
' none it checks everything, from all it clears. Writes the state and
' moves the keyboard cursor there. True when anything toggled; the
' caller runs the handlers.
Private Function ToggleCheckRow(ByVal hostApp As ReDimUI, ByVal checkRow As Long) As Boolean
    Dim rowNumber As Long

    EnsureCheckedSet
    If ItemCount = 0 Then Exit Function
    If checkRow = 0 Then
        If Not mSelectAllRow Then Exit Function
        If mCheckedSet.Count = ItemCount Then
            Set mCheckedSet = New Collection
        Else
            Set mCheckedSet = New Collection
            For rowNumber = 1 To ItemCount
                mCheckedSet.Add rowNumber
            Next rowNumber
        End If
    Else
        If checkRow < 1 Or checkRow > ItemCount Then Exit Function
        TransferSelectionToggle mCheckedSet, checkRow
    End If
    PlaceKeyCursor 0, checkRow
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, CheckedJoined()
    End If
    Touch
    ToggleCheckRow = True
End Function

Private Sub PlaceKeyCursor(ByVal panelIndex As Long, ByVal rowNumber As Long)
    mKeyPanel = panelIndex
    mKeyRow = rowNumber
    mKeyCursorSet = True
End Sub

' Up and Down move a cursor through the focused panel's rows, Left and
' Right switch panels, Space toggles the cursor's row in the selection,
' and Enter moves the panel's selection across, or the cursor's row
' when nothing is selected.
Private Sub HandleTransferKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim rowTotal As Long
    Dim aimRow As Long
    Dim rowsShown As Long

    EnsureTransferCollections
    rowsShown = TransferRowCapacity()
    aimRow = mKeyRow
    Select Case keyText
        Case "{UP}"
            aimRow = mKeyRow - 1
        Case "{DOWN}"
            aimRow = mKeyRow + 1
        Case "{HOME}"
            aimRow = 1
        Case "{END}"
            aimRow = TransferPanelCount(mKeyPanel)
        Case "{PGUP}"
            aimRow = mKeyRow - rowsShown
        Case "{PGDN}"
            aimRow = mKeyRow + rowsShown
        Case "{LEFT}"
            mKeyPanel = 0
        Case "{RIGHT}"
            mKeyPanel = 1
        Case " "
            If mKeyRow >= 1 And mKeyRow <= TransferPanelCount(mKeyPanel) Then
                If mKeyPanel = 0 Then
                    TransferSelectionToggle mSelLeft, mKeyRow
                Else
                    TransferSelectionToggle mSelRight, mKeyRow
                End If
                Touch
            End If
            Exit Sub
        Case "{ENTER}"
            If mKeyRow >= 1 And mKeyRow <= TransferPanelCount(mKeyPanel) Then
                If mKeyPanel = 0 Then
                    If mSelLeft.Count = 0 Then mSelLeft.Add mKeyRow
                Else
                    If mSelRight.Count = 0 Then mSelRight.Add mKeyRow
                End If
                TransferSelected hostApp, (mKeyPanel = 0)
            End If
        Case Else
            Exit Sub
    End Select
    rowTotal = TransferPanelCount(mKeyPanel)
    If aimRow > rowTotal Then aimRow = rowTotal
    If aimRow < 1 Then aimRow = IIf(rowTotal > 0, 1, 0)
    PlaceKeyCursor mKeyPanel, aimRow
    ' The panel window follows the cursor.
    If mKeyPanel = 0 Then
        If aimRow > mLeftOffset + rowsShown Then mLeftOffset = aimRow - rowsShown
        If aimRow >= 1 And aimRow <= mLeftOffset Then mLeftOffset = aimRow - 1
    Else
        If aimRow > mRightOffset + rowsShown Then mRightOffset = aimRow - rowsShown
        If aimRow >= 1 And aimRow <= mRightOffset Then mRightOffset = aimRow - 1
    End If
    Touch
End Sub

Private Function TransferPanelCount(ByVal panelIndex As Long) As Long
    EnsureTransferCollections
    If panelIndex = 0 Then
        TransferPanelCount = mItems.Count
    Else
        TransferPanelCount = mChosenItems.Count
    End If
End Function

' Clearable's button: a small muted circle with an x inside the face's
' right end (left of a combo's caret) while the field holds text.
Private Sub ApplyClearButton( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim clearShape As Shape
    Dim clearKey As String

    If ClearReserve() > 0 Then
        clearKey = PartsState(themeValue)
        If clearKey = mAppliedClearSig Then Exit Sub
        mAppliedClearSig = clearKey
        Set clearShape = EnsurePartShape(hostSheet, "cx", msoShapeOval)
        clearShape.Width = ClearButtonSize()
        clearShape.Height = ClearButtonSize()
        clearShape.Left = mLeft + mWidth - ClearReserve() - _
            IIf(mKind = KIND_COMBO, 16, 0) + 2
        clearShape.Top = mTop + (mHeight - ClearButtonSize()) / 2
        clearShape.Fill.Visible = msoTrue
        clearShape.Fill.ForeColor.RGB = themeValue.MutedColor
        clearShape.Line.Visible = msoFalse
        StyleGlyphText clearShape, ChrW(&H2715), themeValue.OnSurfaceColor, _
            ClearButtonSize() * 0.5, themeValue
        clearShape.Visible = msoTrue
        clearShape.OnAction = mAppliedOnAction
        mainShape.TextFrame2.MarginRight = 6 + ClearReserve() + _
            IIf(mKind = KIND_COMBO, 16, 0)
    ElseIf LenB(mAppliedClearSig) > 0 Then
        mAppliedClearSig = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "cx").Delete
        On Error GoTo 0
        mainShape.TextFrame2.MarginRight = 6
    End If
End Sub

' Under a float field: Validates' message in danger ink at the left, and
' with MaxLength the length against the cap at the right, both in a
' size below the field's.
Private Sub ApplyFieldNotes(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim noteSize As Double
    Dim noteKey As String

    noteSize = mFontSize
    If noteSize <= 0 Then noteSize = themeValue.BaseFontSize
    noteSize = noteSize - 2
    If LenB(mValidationError) > 0 And mVisible Then
        noteKey = PartsState(themeValue) & ":" & mValidationError
    End If
    SyncFieldNote hostSheet, themeValue, "me", noteKey, mValidationError, _
        themeValue.DangerColor, msoAlignLeft, noteSize, mAppliedMessageSig
    noteKey = vbNullString
    If mMaxLength > 0 And mVisible Then
        noteKey = PartsState(themeValue) & ":" & Len(mText) & "/" & mMaxLength
    End If
    SyncFieldNote hostSheet, themeValue, "mc", noteKey, _
        Len(mText) & "/" & mMaxLength, themeValue.OnMutedColor, msoAlignRight, _
        noteSize, mAppliedCountSig
End Sub

' One note under the field; an empty key removes it.
Private Sub SyncFieldNote( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal noteKey As String, _
    ByVal noteWords As String, _
    ByVal inkColor As Long, _
    ByVal textAlign As Long, _
    ByVal noteSize As Double, _
    ByRef appliedKey As String _
)
    Dim noteShape As Shape

    If noteKey = appliedKey Then Exit Sub
    If LenB(noteKey) = 0 Then
        appliedKey = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & partName).Delete
        On Error GoTo 0
        Exit Sub
    End If
    appliedKey = noteKey
    Set noteShape = EnsurePartShape(hostSheet, partName, msoShapeRectangle)
    noteShape.Left = mLeft
    noteShape.Top = mTop + mHeight + 1
    noteShape.Width = mWidth
    noteShape.Height = noteSize + 6
    noteShape.Fill.Visible = msoFalse
    noteShape.Line.Visible = msoFalse
    With noteShape.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorTop
        .WordWrap = msoFalse
        .TextRange.Text = noteWords
        .TextRange.ParagraphFormat.Alignment = textAlign
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = noteSize
            .Bold = msoFalse
            .Fill.ForeColor.RGB = inkColor
        End With
    End With
    noteShape.Visible = msoTrue
    noteShape.OnAction = mAppliedOnAction
End Sub

' Points the clear button takes from the text, 0 while it is not shown.
Private Function ClearReserve() As Double
    If mClearable And mVisible And mEnabled And LenB(mText) > 0 Then
        ClearReserve = ClearButtonSize() + 4
    End If
End Function

Private Function ClearButtonSize() As Double
    ClearButtonSize = mHeight - 4
    If ClearButtonSize > 16 Then ClearButtonSize = 16
End Function

' A click in a float field: the first focuses it (a combo also opens its
' list), and every click puts the caret at the character under the
' pointer. A second click on the same spot inside the Windows
' double-click time selects the word there.
Private Sub ClickField(ByVal hostApp As ReDimUI)
    Dim wasFocused As Boolean
    Dim hitPos As Long
    Dim nowMs As LongLong

    wasFocused = ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId)
    hitPos = PointerCaretIndex(wasFocused)
    If Not wasFocused Then FocusField hostApp
    nowMs = GetTickCount64()
    If hitPos >= 0 Then
        If wasFocused And Abs(hitPos - mLastFieldClickPos) <= 1 _
            And nowMs - mLastFieldClickMs <= GetDoubleClickTime() Then
            SelectWordAt hitPos
        Else
            mCaretPos = hitPos
            mSelAnchor = hitPos
        End If
    End If
    mLastFieldClickMs = nowMs
    mLastFieldClickPos = hitPos
    mLastEditKind = vbNullString
    mCaretOn = True
    mCaretAccumMs = 0
    Touch
End Sub

' The caret index the pointer points at, read from the face as drawn:
' the nearest character on the nearest line, before it or after it by
' which half the pointer is over. A focused face maps back through
' mFaceMap; an unfocused face is the text itself. -1 when unreadable,
' or when the pointer is off the face (a keyboard or scripted stand-in
' for the click, or a pointer that moved on), so the caret stays put.
Private Function PointerCaretIndex(ByVal focusedFace As Boolean) As Long
    Dim pointsX As Double
    Dim pointsY As Double
    Dim faceRange As Object
    Dim charCount As Long
    Dim idx As Long
    Dim boxLeft As Double
    Dim boxTop As Double
    Dim boxWidth As Double
    Dim boxHeight As Double
    Dim lineGap As Double
    Dim sideGap As Double
    Dim distance As Double
    Dim bestDistance As Double
    Dim bestChar As Long
    Dim afterChar As Boolean
    Dim textChar As Long

    PointerCaretIndex = -1
    If LenB(mText) = 0 Then
        PointerCaretIndex = 0
        Exit Function
    End If
    If Not PointerPoints(pointsX, pointsY) Then Exit Function
    If pointsX < mLeft - 4 Or pointsX > mLeft + mWidth + 4 _
        Or pointsY < mTop - 4 Or pointsY > mTop + mHeight + 4 Then
        Exit Function
    End If
    On Error GoTo Unreadable
    Set faceRange = OwnerApp.Sheet.Shapes(ShapeName).TextFrame2.TextRange
    charCount = faceRange.Length
    If charCount = 0 Then
        PointerCaretIndex = 0
        Exit Function
    End If
    bestDistance = 1E+30
    For idx = 1 To charCount
        With faceRange.Characters(idx, 1)
            boxLeft = .BoundLeft
            boxTop = .BoundTop
            boxWidth = .BoundWidth
            boxHeight = .BoundHeight
        End With
        lineGap = 0
        If pointsY < boxTop Then lineGap = boxTop - pointsY
        If pointsY > boxTop + boxHeight Then lineGap = pointsY - boxTop - boxHeight
        sideGap = 0
        If pointsX < boxLeft Then sideGap = boxLeft - pointsX
        If pointsX > boxLeft + boxWidth Then sideGap = pointsX - boxLeft - boxWidth
        ' The line decides first, the position along it second.
        distance = lineGap * 10000 + sideGap
        If distance < bestDistance Then
            bestDistance = distance
            bestChar = idx
            afterChar = (pointsX > boxLeft + boxWidth / 2)
        End If
    Next idx
    If Not focusedFace Then
        textChar = bestChar
        If textChar > Len(mText) Then textChar = Len(mText)
        PointerCaretIndex = IIf(afterChar, textChar, textChar - 1)
        Exit Function
    End If
    If bestChar > mFaceLength Then bestChar = mFaceLength
    textChar = mFaceMap(bestChar)
    If textChar > 0 Then
        PointerCaretIndex = IIf(afterChar, textChar, textChar - 1)
        Exit Function
    End If
    ' The bar or an ellipsis: the caret goes next to the nearest real
    ' character on the side the pointer is.
    For idx = bestChar - 1 To 1 Step -1
        If mFaceMap(idx) > 0 Then
            PointerCaretIndex = mFaceMap(idx)
            Exit Function
        End If
    Next idx
    For idx = bestChar + 1 To mFaceLength
        If mFaceMap(idx) > 0 Then
            PointerCaretIndex = mFaceMap(idx) - 1
            Exit Function
        End If
    Next idx
    PointerCaretIndex = 0
    Exit Function
Unreadable:
    PointerCaretIndex = -1
End Function

' Selects the word at or next to a caret index, as a double click does.
Private Sub SelectWordAt(ByVal atPos As Long)
    Dim wordFrom As Long
    Dim wordTo As Long

    wordFrom = atPos
    wordTo = atPos
    Do While wordFrom > 0
        If Not IsWordChar(Mid$(mText, wordFrom, 1)) Then Exit Do
        wordFrom = wordFrom - 1
    Loop
    Do While wordTo < Len(mText)
        If Not IsWordChar(Mid$(mText, wordTo + 1, 1)) Then Exit Do
        wordTo = wordTo + 1
    Loop
    mSelAnchor = wordFrom
    mCaretPos = wordTo
End Sub

' After any edit: OnInput now or once typing pauses, and a recheck while
' a validation message shows.
Private Sub NoteFieldEdit(ByVal hostApp As ReDimUI)
    If mValidationLive Then RunValidation hostApp
    If LenB(mInputProc) = 0 Then Exit Sub
    If mDebounceMs <= 0 Then
        FireInput hostApp
    Else
        mInputDueMs = GetTickCount64() + mDebounceMs
    End If
End Sub

Private Sub FireInput(ByVal hostApp As ReDimUI)
    mInputDueMs = 0
    If LenB(mInputProc) = 0 Then Exit Sub
    ReDimUI.RunHandlerAs hostApp, Me, mInputProc
End Sub

' Runs Validates' check on the text. A check that fails to run counts
' as a pass and goes to the app's error handler.
Private Sub RunValidation(ByVal hostApp As ReDimUI)
    Dim verdict As Variant

    If LenB(mValidateProc) = 0 Then Exit Sub
    On Error GoTo CheckFailed
    verdict = Application.Run(mValidateProc, mText)
    On Error GoTo 0
    mValidationError = CStr(verdict)
    mValidationLive = (LenB(mValidationError) > 0)
    Exit Sub
CheckFailed:
    mValidationError = vbNullString
    mValidationLive = False
    hostApp.ReportAppError "ReDim validation error " & Err.Number & ": " & Err.Description
End Sub

' The clear button: empties the field and leaves it focused for typing;
' a combo's list opens unfiltered, as after deleting the text by hand.
Private Sub ClearFieldText(ByVal hostApp As ReDimUI)
    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then
        FocusField hostApp, False
    End If
    mText = vbNullString
    mCaretPos = 0
    If mKind = KIND_COMBO Then
        mValue = 0
        If Not mSelectOpen Then NoteListOpened
        mSelectOpen = True
        mFilterPristine = False
        mListOffset = 0
        mListHighlight = 0
    End If
    NoteFieldEdit hostApp
    Touch
End Sub

' A focused control that is not a text field wears a ring: an accent
' outline three points outside its bounds. Float fields show focus on
' their own border. A focused control that turns hidden or disabled
' gives focus up here.
Private Sub ApplyFocusRing(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim ringOn As Boolean
    Dim ringShape As Shape
    Dim ringKey As String

    ringOn = ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId)
    If ringOn And Not IsFocusable() Then
        ReDimUI.ClearKeyboardFocus
        ringOn = False
    End If
    If IsFloatField() Then Exit Sub
    If ringOn Then
        ringKey = PartsState(themeValue)
        If ringKey = mAppliedRingSig Then Exit Sub
        mAppliedRingSig = ringKey
        Set ringShape = EnsurePartShape(hostSheet, "fr", msoShapeRoundedRectangle)
        On Error Resume Next
        ringShape.Adjustments(1) = 0.12
        On Error GoTo 0
        ringShape.Left = mLeft - 3
        ringShape.Top = mTop - 3
        ringShape.Width = mWidth + 6
        ringShape.Height = mHeight + 6
        ringShape.Fill.Visible = msoFalse
        ringShape.Line.Visible = msoTrue
        ringShape.Line.ForeColor.RGB = themeValue.PrimaryColor
        ringShape.Line.Weight = 2
        ringShape.Visible = msoTrue
        ringShape.OnAction = mAppliedOnAction
    ElseIf LenB(mAppliedRingSig) > 0 Then
        mAppliedRingSig = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "fr").Delete
        On Error GoTo 0
    End If
End Sub

' The row a focused list's arrow keys have reached: a dotted outline in
' surface ink over the row, readable on the selection fill as well.
Private Sub SyncKeyCursor( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal showIt As Boolean, _
    ByVal cursorLeft As Double, _
    ByVal cursorTop As Double, _
    ByVal cursorWidth As Double, _
    ByVal cursorHeight As Double _
)
    Dim cursorShape As Shape
    Dim cursorKey As String

    If showIt Then
        cursorKey = PartsState(themeValue) & ":" & CStr(cursorLeft) & ":" & _
            CStr(cursorTop) & ":" & CStr(cursorWidth) & ":" & CStr(cursorHeight)
        If cursorKey = mAppliedKeyCursorSig Then Exit Sub
        mAppliedKeyCursorSig = cursorKey
        Set cursorShape = EnsurePartShape(hostSheet, "kr", msoShapeRectangle)
        cursorShape.Left = cursorLeft
        cursorShape.Top = cursorTop
        cursorShape.Width = cursorWidth
        cursorShape.Height = cursorHeight
        cursorShape.Fill.Visible = msoFalse
        cursorShape.Line.Visible = msoTrue
        cursorShape.Line.ForeColor.RGB = themeValue.OnSurfaceColor
        cursorShape.Line.Weight = 1.25
        cursorShape.Line.DashStyle = msoLineRoundDot
        cursorShape.Visible = msoTrue
        cursorShape.OnAction = mAppliedOnAction
        cursorShape.ZOrder msoBringToFront
    ElseIf LenB(mAppliedKeyCursorSig) > 0 Then
        mAppliedKeyCursorSig = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "kr").Delete
        On Error GoTo 0
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
            ApplyToastParts mainShape, hostSheet, themeValue
    End Select
End Sub

' A toast's tone from its variant: none unless a variant builder ran;
' Primary reads as information.
Private Function ToastIcon() As String
    If Not mVariantChosen Then Exit Function
    Select Case mStyleVariant
        Case VARIANT_PRIMARY
            ToastIcon = ChrW(&H2139)
        Case VARIANT_SUCCESS
            ToastIcon = ChrW(&H2713)
        Case VARIANT_WARNING
            ToastIcon = ChrW(&H26A0)
        Case VARIANT_DANGER
            ToastIcon = ChrW(&H2715)
    End Select
End Function

Private Function ToastToneColor(ByVal themeValue As ReDimUI) As Long
    Select Case mStyleVariant
        Case VARIANT_SUCCESS
            ToastToneColor = themeValue.SuccessColor
        Case VARIANT_WARNING
            ToastToneColor = themeValue.WarningColor
        Case VARIANT_DANGER
            ToastToneColor = themeValue.DangerColor
        Case Else
            ToastToneColor = themeValue.PrimaryColor
    End Select
End Function

' The tone's icon color and edge, the close button, and the action
' button, placed at the toast's current visual top; the pump carries
' them through the slide and the fade from there.
Private Sub ApplyToastParts( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim visualTop As Double
    Dim closeButton As Shape
    Dim actionButton As Shape
    Dim actionWidth As Double
    Dim resolvedSize As Double

    visualTop = mTop + mSlideRemaining
    If ExtrasCurrent(PartsState(themeValue) & ":" & ToastIcon() & ":" & _
        mToastActionLabel & ":" & CStr(visualTop)) Then Exit Sub
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    If LenB(ToastIcon()) > 0 Then
        mainShape.Line.ForeColor.RGB = ToastToneColor(themeValue)
        mainShape.Line.Weight = 1.5
        With mainShape.TextFrame2.TextRange
            .Characters(1, 1).Font.Fill.ForeColor.RGB = ToastToneColor(themeValue)
            With .ParagraphFormat.TabStops
                If .Count = 0 Then .Add msoTabStopLeft, 18
            End With
        End With
    End If

    Set closeButton = EnsurePartShape(hostSheet, "tx", msoShapeRectangle)
    closeButton.Left = mLeft + mWidth - PAGER_TARGET - 3
    closeButton.Top = visualTop + 3
    closeButton.Width = PAGER_TARGET
    closeButton.Height = PAGER_TARGET
    closeButton.Fill.Visible = msoTrue
    closeButton.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
    closeButton.Line.Visible = msoFalse
    StyleGlyphText closeButton, ChrW(&H2715), themeValue.OnMutedColor, 8, _
        themeValue
    closeButton.Visible = BoolToMso(mVisible)
    closeButton.OnAction = mAppliedOnAction

    If LenB(mToastActionLabel) > 0 Then
        ' Sized from the label; ink in surface text, since the accent
        ' does not reach 4.5:1 as text on every theme's surface.
        actionWidth = Len(mToastActionLabel) * resolvedSize * 0.62 + 14
        Set actionButton = EnsurePartShape( _
            hostSheet, "ta", msoShapeRoundedRectangle)
        actionButton.Left = mLeft + mWidth - PAGER_TARGET - 7 - actionWidth
        actionButton.Top = visualTop + (mHeight - 22) / 2
        actionButton.Width = actionWidth
        actionButton.Height = 22
        actionButton.Fill.Visible = msoTrue
        actionButton.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
        actionButton.Line.Visible = msoTrue
        actionButton.Line.ForeColor.RGB = themeValue.PrimaryColor
        actionButton.Line.Weight = 1
        StyleGlyphText actionButton, mToastActionLabel, _
            themeValue.OnSurfaceColor, resolvedSize - 1, themeValue
        actionButton.Visible = BoolToMso(mVisible)
        actionButton.OnAction = mAppliedOnAction
        mainShape.TextFrame2.MarginRight = PAGER_TARGET + 11 + actionWidth
    Else
        mainShape.TextFrame2.MarginRight = PAGER_TARGET + 6
    End If
    mainShape.TextFrame2.MarginLeft = 8
End Sub

' Carries the close and action buttons with the toast's slide and fade.
Private Sub TrackToastParts(ByVal hostSheet As Worksheet)
    Dim visualTop As Double

    visualTop = mTop + mSlideRemaining
    On Error Resume Next
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "tx").Top = visualTop + 3
    hostSheet.Shapes(ShapeName & PART_SEPARATOR & "ta").Top = _
        visualTop + (mHeight - 22) / 2
    If mToastExiting Then
        FadeToastPart hostSheet, "tx"
        FadeToastPart hostSheet, "ta"
    End If
    On Error GoTo 0
End Sub

Private Sub FadeToastPart(ByVal hostSheet As Worksheet, ByVal partName As String)
    On Error Resume Next
    With hostSheet.Shapes(ShapeName & PART_SEPARATOR & partName)
        .Fill.Transparency = 1 - mExitRemaining
        .Line.Transparency = 1 - mExitRemaining
        .TextFrame2.TextRange.Font.Fill.Transparency = 1 - mExitRemaining
    End With
    On Error GoTo 0
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

' The state every composite part depends on: theme, geometry, visibility,
' enablement, and font. Kinds append their own state to it.
Private Function PartsState(ByVal themeValue As ReDimUI) As String
    PartsState = CStr(ObjPtr(themeValue)) & ":" & CStr(mLeft) & ":" & _
        CStr(mTop) & ":" & CStr(mWidth) & ":" & CStr(mHeight) & ":" & _
        CStr(mVisible) & ":" & CStr(mEnabled) & ":" & CStr(mFontSize) & _
        ":" & CStr(mFontBold) & ":" & mAppliedOnAction
End Function

' True when the kind's extra parts were last written for exactly this
' state, so every part write can be skipped; otherwise records the state
' and returns False. Render forgets it, so an explicit render rewrites.
Private Function ExtrasCurrent(ByVal partsKey As String) As Boolean
    If mAppliedValid And partsKey = mAppliedExtrasSig Then
        ExtrasCurrent = True
    Else
        mAppliedExtrasSig = partsKey
    End If
End Function

Friend Sub ForgetExtrasSignature()
    mAppliedExtrasSig = vbNullString
End Sub

' Every part key dropped, so the next apply rewrites every part: for a
' main shape recreated from scratch and for teardown.
Friend Sub ForgetPartSignatures()
    mAppliedExtrasSig = vbNullString
    mAppliedListSig = vbNullString
    mAppliedProgressSig = vbNullString
    mAppliedTransferSig = vbNullString
    mAppliedCheckSig = vbNullString
    mAppliedCheckHeaderSig = vbNullString
    ' Parts that may or may not exist read as unknown, which forces one
    ' redraw when needed and one delete when not.
    mAppliedPagerUpSig = "?"
    mAppliedPagerDownSig = "?"
    mAppliedEmptyRowSig = "?"
    mAppliedHeaderLeftSig = vbNullString
    mAppliedHeaderRightSig = vbNullString
    mAppliedScrollLeftSig = "?"
    mAppliedScrollRightSig = "?"
    ' Focus and clear parts start absent: probing for them on every new
    ' control's first draw cost a failed lookup each, for parts only a
    ' focused or typed-in control ever has.
    mAppliedRingSig = vbNullString
    mAppliedKeyCursorSig = vbNullString
    mAppliedClearSig = vbNullString
    Erase mAppliedRowStates
    Erase mAppliedLeftRowStates
    Erase mAppliedRightRowStates
End Sub

' Box at the left edge, check glyph inside it when checked, caption filling
' the remainder. Both pieces click-toggle through the dispatcher.
Private Sub ApplyTickBox( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim captionShape As Shape
    Dim boxSize As Double

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mChecked) & _
        ":" & mText & ":" & mAccessKey) Then Exit Sub
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
    If LenB(mAccessKey) > 0 Or LenB(mAppliedAccessKey) > 0 Then
        MarkAccessKey captionShape, mText
    End If
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
    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mValue) & _
        JoinedSig(mItems)) Then Exit Sub
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
' The layout key covers what every row shares; when it changes, every
' row is rewritten. Otherwise a row whose check changed rewrites its box
' state and one whose caption changed rewrites its text, so a toggle
' costs one box and the header instead of the list.
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
    Dim checkedTotal As Long
    Dim layoutKey As String
    Dim rowKey As String
    Dim priorKey As String
    Dim headerKey As String
    Dim headerText As String
    Dim boxCreated As Boolean
    Dim captionCreated As Boolean
    Dim fullPass As Boolean
    Dim checkedFlags() As Boolean
    Dim entry As Variant

    EnsureCheckedSet
    If Not mItems Is Nothing Then rowTotal = mItems.Count
    If rowTotal = 0 Then
        SyncKeyCursor hostSheet, themeValue, False, 0, 0, 0, 0
        Exit Sub
    End If
    PruneTransferSelection mCheckedSet, rowTotal
    layoutKey = PartsState(themeValue) & ":" & CStr(mSelectAllRow) & _
        ":" & CStr(rowTotal)
    fullPass = Not (mAppliedValid And layoutKey = mAppliedCheckSig)
    If fullPass Then
        mAppliedCheckSig = layoutKey
        mAppliedCheckHeaderSig = vbNullString
        ReDim mAppliedRowStates(1 To rowTotal)
    End If
    ReDim checkedFlags(1 To rowTotal)
    For Each entry In mCheckedSet
        checkedFlags(CLng(entry)) = True
    Next entry
    checkedTotal = mCheckedSet.Count
    headerRows = 0
    If mSelectAllRow Then headerRows = 1
    rowPitch = mHeight / (rowTotal + headerRows)
    boxSize = rowPitch - 6
    If boxSize > 14 Then boxSize = 14
    If boxSize < 10 Then boxSize = 10

    If mSelectAllRow Then
        headerKey = CStr(checkedTotal) & "/" & CStr(rowTotal)
        priorKey = mAppliedCheckHeaderSig
        If headerKey <> priorKey Then
            mAppliedCheckHeaderSig = headerKey
            Set box = EnsurePartShape( _
                hostSheet, "mb", msoShapeRoundedRectangle, boxCreated)
            Set captionShape = EnsurePartShape( _
                hostSheet, "mt", msoShapeRectangle, captionCreated)
            headerText = "Select all (" & checkedTotal & "/" & rowTotal & ")"
            If LenB(priorKey) > 0 And Not (boxCreated Or captionCreated) Then
                StyleCheckBoxState box, boxSize, themeValue, _
                    (checkedTotal = rowTotal), (checkedTotal > 0)
                captionShape.TextFrame2.TextRange.Text = headerText
            Else
                On Error Resume Next
                box.Adjustments(1) = 0.2
                On Error GoTo 0
                StyleCheckBoxFace box, boxSize, mTop, rowPitch, themeValue, _
                    (checkedTotal = rowTotal), (checkedTotal > 0)
                captionShape.Left = mLeft + boxSize + 8
                captionShape.Top = mTop
                captionShape.Width = mWidth - boxSize - 8
                captionShape.Height = rowPitch
                StyleCaptionPart captionShape, headerText, themeValue
                captionShape.Visible = BoolToMso(mVisible)
                captionShape.OnAction = mAppliedOnAction
            End If
        End If
    ElseIf fullPass Then
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "mb").Delete
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "mt").Delete
        On Error GoTo 0
    End If

    For idx = 1 To rowTotal
        rowKey = IIf(checkedFlags(idx), "1", "0") & ItemAt(idx)
        priorKey = mAppliedRowStates(idx)
        If rowKey <> priorKey Then
            mAppliedRowStates(idx) = rowKey
            rowTop = mTop + (idx - 1 + headerRows) * rowPitch
            boxCreated = False
            If idx = 1 Then
                Set box = mainShape
            Else
                Set box = EnsurePartShape( _
                    hostSheet, "b" & idx, msoShapeRoundedRectangle, boxCreated)
            End If
            Set captionShape = EnsurePartShape( _
                hostSheet, "t" & idx, msoShapeRectangle, captionCreated)
            If LenB(priorKey) > 0 And Not (boxCreated Or captionCreated) Then
                ' Drawn under this layout: rewrite only the box state or the
                ' caption text, whichever moved.
                If Left$(priorKey, 1) <> Left$(rowKey, 1) Then
                    StyleCheckBoxState box, boxSize, themeValue, _
                        checkedFlags(idx), checkedFlags(idx)
                End If
                If Mid$(priorKey, 2) <> Mid$(rowKey, 2) Then
                    captionShape.TextFrame2.TextRange.Text = ItemAt(idx)
                End If
            Else
                If idx > 1 Then
                    On Error Resume Next
                    box.Adjustments(1) = 0.2
                    On Error GoTo 0
                End If
                StyleCheckBoxFace box, boxSize, rowTop, rowPitch, themeValue, _
                    checkedFlags(idx), checkedFlags(idx)
                captionShape.Left = mLeft + boxSize + 8
                captionShape.Top = rowTop
                captionShape.Width = mWidth - boxSize - 8
                captionShape.Height = rowPitch
                StyleCaptionPart captionShape, ItemAt(idx), themeValue
                captionShape.Visible = BoolToMso(mVisible)
                captionShape.OnAction = mAppliedOnAction
            End If
        End If
    Next idx
    If fullPass Then
        For idx = rowTotal + 1 To mAppliedOptionCount
            On Error Resume Next
            hostSheet.Shapes(ShapeName & PART_SEPARATOR & "b" & idx).Delete
            hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & idx).Delete
            On Error GoTo 0
        Next idx
        mAppliedOptionCount = rowTotal
    End If
    If mKeyRow > rowTotal Then mKeyRow = rowTotal
    If mKeyRow < 1 - headerRows Then mKeyRow = 1
    SyncKeyCursor hostSheet, themeValue, _
        ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) And mVisible, _
        mLeft - 2, mTop + (mKeyRow - 1 + headerRows) * rowPitch, _
        mWidth + 4, rowPitch
End Sub

' The parts of a checkbox face that follow its state: fill, border, and
' glyph text, size, and ink. Geometry and glyph layout stay as
' StyleCheckBoxFace left them, since shape text keeps its formatting
' across a change to or from empty text.
Private Sub StyleCheckBoxState( _
    ByVal box As Shape, _
    ByVal boxSize As Double, _
    ByVal themeValue As ReDimUI, _
    ByVal fullState As Boolean, _
    ByVal partialState As Boolean _
)
    If partialState And mEnabled Then
        box.Fill.ForeColor.RGB = ResolveFill(themeValue)
        box.Line.Visible = msoFalse
    Else
        box.Fill.ForeColor.RGB = themeValue.SurfaceColor
        box.Line.Visible = msoTrue
        box.Line.ForeColor.RGB = themeValue.BorderColor
        box.Line.Weight = 1.25
    End If
    With box.TextFrame2.TextRange
        If fullState Then
            .Text = ChrW(10003)
            .Font.Size = boxSize * 0.62
        ElseIf partialState Then
            .Text = "-"
            .Font.Size = boxSize * 0.8
        Else
            .Text = vbNullString
        End If
        If partialState Then
            .Font.Fill.ForeColor.RGB = _
                IIf(mEnabled, themeValue.OnPrimaryColor, themeValue.OnMutedColor)
        End If
    End With
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

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mValue)) Then Exit Sub
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

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mChecked)) Then Exit Sub
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
    ' On, the knob takes the ink that reads on the accent track, the way
    ' Windows draws it: white on the light theme's green, dark on the
    ' dark theme's, black on high contrast's yellow.
    If mChecked And mEnabled Then
        knob.Fill.ForeColor.RGB = themeValue.OnPrimaryColor
    Else
        knob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
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
' The frame (both panels and the move buttons) is rewritten only when the
' layout changes; headers, rows, and paging arrows each keep the state
' they were last drawn with, so a row click rewrites that row alone.
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
    Dim layoutKey As String

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
    panelW = (mWidth - BTN_W - 4 * PANEL_GAP) / 2
    If panelW < 40 Then panelW = 40
    panelTop = mTop + HEADER_H
    panelH = mHeight - HEADER_H
    If panelH < 30 Then panelH = 30
    rightX = mLeft + panelW + BTN_W + 4 * PANEL_GAP
    btnX = mLeft + panelW + 2 * PANEL_GAP

    layoutKey = PartsState(themeValue)
    If Not (mAppliedValid And layoutKey = mAppliedTransferSig) Then
        mAppliedTransferSig = layoutKey
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
    End If

    ApplyTransferHeader hostSheet, themeValue, "hl", mLeft, panelW, HEADER_H, _
        IIf(LenB(mCaptionLeft) > 0, mCaptionLeft, "Available") & _
        " (" & mItems.Count & ")", mAppliedHeaderLeftSig
    ApplyTransferHeader hostSheet, themeValue, "hr", rightX, panelW, HEADER_H, _
        IIf(LenB(mCaptionRight) > 0, mCaptionRight, "Selected") & _
        " (" & mChosenItems.Count & ")", mAppliedHeaderRightSig
    SyncTransferRows hostSheet, themeValue, mItems, "al", _
        mLeft, panelTop, panelW, panelH, mSelLeft, mLeftOffset, _
        mAppliedLeftRows, mAppliedLeftRowStates
    SyncTransferRows hostSheet, themeValue, mChosenItems, "cl", _
        rightX, panelTop, panelW, panelH, mSelRight, mRightOffset, _
        mAppliedRightRows, mAppliedRightRowStates
    ApplyTransferScroller hostSheet, themeValue, "alu", "ald", _
        mLeft, panelTop, panelW, panelH, mItems.Count, mAppliedScrollLeftSig
    ApplyTransferScroller hostSheet, themeValue, "clu", "cld", _
        rightX, panelTop, panelW, panelH, mChosenItems.Count, _
        mAppliedScrollRightSig

    ' The keyboard cursor, while the list holds focus.
    Dim cursorItems As Long
    Dim cursorRow As Long
    Dim cursorWidth As Double
    cursorItems = TransferPanelCount(mKeyPanel)
    If mKeyRow > cursorItems Then mKeyRow = cursorItems
    cursorRow = mKeyRow - IIf(mKeyPanel = 0, mLeftOffset, mRightOffset)
    cursorWidth = panelW - 6
    If cursorItems > TransferRowCapacity() Then
        cursorWidth = cursorWidth - PAGER_TARGET - 2
    End If
    SyncKeyCursor hostSheet, themeValue, _
        ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) And mVisible _
            And mKeyRow >= 1 And cursorRow >= 1 _
            And cursorRow <= TransferRowCapacity(), _
        IIf(mKeyPanel = 0, mLeft, rightX) + 3, _
        panelTop + 3 + (cursorRow - 1) * 20, cursorWidth, 20
End Sub

' A panel's counted header, rewritten only when its text or the layout
' changes.
Private Sub ApplyTransferHeader( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal headerLeft As Double, _
    ByVal headerWidth As Double, _
    ByVal headerHeight As Double, _
    ByVal headerText As String, _
    ByRef appliedKey As String _
)
    Dim headerShape As Shape
    Dim headerKey As String

    headerKey = PartsState(themeValue) & ":" & headerText
    If headerKey = appliedKey Then Exit Sub
    appliedKey = headerKey
    Set headerShape = EnsurePartShape(hostSheet, partName, msoShapeRectangle)
    headerShape.Left = headerLeft
    headerShape.Top = mTop - 1
    headerShape.Width = headerWidth
    headerShape.Height = headerHeight
    StyleCaptionPart headerShape, headerText, themeValue
    headerShape.Visible = BoolToMso(mVisible)
End Sub

' Small paging arrows on a panel's right edge, shown only when the list
' outgrows the panel. appliedKey follows SyncListPager's convention:
' empty when absent, "?" when unknown after Render.
Private Sub ApplyTransferScroller( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal upPart As String, _
    ByVal downPart As String, _
    ByVal panelX As Double, _
    ByVal panelTop As Double, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal itemTotal As Long, _
    ByRef appliedKey As String _
)
    Dim scrollButton As Shape
    Dim scrollKey As String
    Dim firstDraw As Boolean

    If itemTotal > TransferRowCapacity() And mVisible Then
        scrollKey = PartsState(themeValue) & ":" & CStr(panelX) & ":" & _
            CStr(panelTop) & ":" & CStr(panelW) & ":" & CStr(panelH)
        If scrollKey = appliedKey Then Exit Sub
        firstDraw = (LenB(appliedKey) = 0 Or appliedKey = "?")
        appliedKey = scrollKey
        Set scrollButton = EnsurePartShape( _
            hostSheet, upPart, msoShapeRoundedRectangle)
        scrollButton.Left = panelX + panelW - PAGER_TARGET - 3
        scrollButton.Top = panelTop + 3
        scrollButton.Width = PAGER_TARGET
        scrollButton.Height = PAGER_TARGET
        scrollButton.Fill.Visible = msoTrue
        scrollButton.Fill.ForeColor.RGB = themeValue.MutedColor
        StyleGlyphText scrollButton, ChrW(9650), _
            ResolveDrawnInk(themeValue), 9, themeValue
        scrollButton.Visible = msoTrue
        scrollButton.OnAction = mAppliedOnAction
        If firstDraw Then scrollButton.ZOrder msoBringToFront
        Set scrollButton = EnsurePartShape( _
            hostSheet, downPart, msoShapeRoundedRectangle)
        scrollButton.Left = panelX + panelW - PAGER_TARGET - 3
        scrollButton.Top = panelTop + panelH - PAGER_TARGET - 3
        scrollButton.Width = PAGER_TARGET
        scrollButton.Height = PAGER_TARGET
        scrollButton.Fill.Visible = msoTrue
        scrollButton.Fill.ForeColor.RGB = themeValue.MutedColor
        StyleGlyphText scrollButton, ChrW(9660), _
            ResolveDrawnInk(themeValue), 9, themeValue
        scrollButton.Visible = msoTrue
        scrollButton.OnAction = mAppliedOnAction
        If firstDraw Then scrollButton.ZOrder msoBringToFront
    ElseIf LenB(appliedKey) > 0 Then
        appliedKey = vbNullString
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

' One panel's window of rows. Each row keeps the state it was last drawn
' with, so selecting a row or paging rewrites only rows that changed; a
' row whose frame held rewrites only its fill, text, and ink, and a row
' is brought to the front only when first drawn.
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
    ByRef appliedRows As Long, _
    ByRef rowStates() As String _
)
    Const ROW_H As Double = 20
    Dim rowShape As Shape
    Dim idx As Long
    Dim wanted As Long
    Dim itemIndex As Long
    Dim rowWidth As Double
    Dim resolvedSize As Double
    Dim rowSelected As Boolean
    Dim rowKey As String
    Dim priorKey As String
    Dim rowText As String
    Dim rowFill As Long
    Dim rowInk As Long
    Dim partCreated As Boolean
    Dim layoutKey As String

    wanted = itemsRef.Count - scrollOffset
    If wanted > Int((panelH - 6) / ROW_H) Then
        wanted = Int((panelH - 6) / ROW_H)
    End If
    If wanted < 0 Then wanted = 0
    If Not mVisible Then wanted = 0
    rowWidth = panelW - 6
    If itemsRef.Count > Int((panelH - 6) / ROW_H) Then
        ' Leave room for the paging arrows.
        rowWidth = rowWidth - PAGER_TARGET - 2
    End If
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    layoutKey = PartsState(themeValue) & ":" & CStr(panelX) & ":" & _
        CStr(panelTop) & ":" & CStr(rowWidth)
    EnsureRowStates rowStates, wanted
    For idx = 1 To wanted
        itemIndex = scrollOffset + idx
        rowSelected = TransferSelectionHas(selectionSet, itemIndex)
        If rowSelected Then
            rowFill = themeValue.PrimaryColor
            rowInk = themeValue.OnPrimaryColor
        Else
            rowFill = ResolveSurfaceFill(themeValue)
            rowInk = themeValue.OnSurfaceColor
        End If
        rowText = GutterText(rowSelected, CStr(itemsRef.Item(itemIndex)))
        rowKey = layoutKey & Chr$(1) & CStr(rowFill) & ":" & _
            CStr(rowInk) & ":" & rowText
        priorKey = rowStates(idx)
        If rowKey <> priorKey Then
            rowStates(idx) = rowKey
            Set rowShape = EnsurePartShape( _
                hostSheet, partPrefix & idx, msoShapeRectangle, partCreated)
            If KeyFrame(priorKey) = layoutKey And Not partCreated Then
                ' Same frame: only the fill, the text, and the ink move.
                rowShape.Fill.ForeColor.RGB = rowFill
                With rowShape.TextFrame2.TextRange
                    .Text = rowText
                    .Font.Fill.ForeColor.RGB = rowInk
                End With
            Else
                rowShape.Left = panelX + 3
                rowShape.Top = panelTop + 3 + (idx - 1) * ROW_H
                rowShape.Width = rowWidth
                rowShape.Height = ROW_H
                rowShape.Fill.Visible = msoTrue
                rowShape.Fill.ForeColor.RGB = rowFill
                With rowShape.TextFrame2
                    .MarginLeft = 6
                    .MarginRight = 4
                    .MarginTop = 1
                    .MarginBottom = 1
                    .VerticalAnchor = msoAnchorMiddle
                    .WordWrap = msoFalse
                    .TextRange.Text = rowText
                    .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                    With .TextRange.Font
                        .Name = themeValue.FontName
                        .Size = resolvedSize
                        .Bold = msoFalse
                        .Fill.ForeColor.RGB = rowInk
                    End With
                End With
                If LenB(priorKey) = 0 Or partCreated Then
                    EnsureGutterTab rowShape
                End If
                rowShape.Visible = msoTrue
                rowShape.OnAction = mAppliedOnAction
                If LenB(priorKey) = 0 Then rowShape.ZOrder msoBringToFront
            End If
        End If
    Next idx
    EnsureRowStates rowStates, appliedRows
    For idx = wanted + 1 To appliedRows
        rowStates(idx) = vbNullString
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
    ApplyListCaret hostSheet, themeValue, 14
    SyncDropList hostSheet, themeValue
End Sub

' The drop-down caret at the face's right edge, shared by the SelectBox
' and the ComboBox.
Private Sub ApplyListCaret( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal rightInset As Double _
)
    Dim caret As Shape

    If ExtrasCurrent(PartsState(themeValue)) Then Exit Sub
    Set caret = EnsurePartShape(hostSheet, "caret", msoShapeIsoscelesTriangle)
    caret.Width = 8
    caret.Height = 5
    caret.Left = mLeft + mWidth - rightInset
    caret.Top = mTop + (mHeight - 5) / 2
    caret.Rotation = 180
    caret.Fill.Visible = msoTrue
    caret.Fill.ForeColor.RGB = themeValue.OnMutedColor
    caret.Visible = BoolToMso(mVisible)
    caret.OnAction = mAppliedOnAction
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
    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(fraction) & _
        ":" & CStr(mSlideTracking)) Then Exit Sub
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
    Dim pointsX As Double
    Dim pointsY As Double

    If Not PointerPoints(pointsX, pointsY) Then Exit Function
    PressClaimedByChrome = hostApp.PointClaimedByChrome( _
        pointsX, pointsY, mComponentId)
End Function

' The pointer in sheet points: the test override when one is set, else
' the cursor mapped through the active window. False when unmappable.
Private Function PointerPoints( _
    ByRef pointsX As Double, _
    ByRef pointsY As Double _
) As Boolean
    Dim screenPoint As RdxScreenPoint

    If ReDimUI.PointerOverride(pointsX, pointsY) Then
        PointerPoints = True
        Exit Function
    End If
    If GetCursorPos(screenPoint) = 0 Then Exit Function
    pointsX = ScreenXToSheetPoints(screenPoint.pixelX)
    If pointsX = UNMAPPABLE_POINTS Then Exit Function
    pointsY = ScreenYToSheetPoints(screenPoint.pixelY)
    If pointsY = UNMAPPABLE_POINTS Then Exit Function
    PointerPoints = True
End Function

' True when a sheet point lies on the face or on the open list's rows and
' pagers, below the face or above it, with a two-point slop.
Private Function ListAreaHolds( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double _
) As Boolean
    If pointsX < mLeft - 2 Or pointsX > mLeft + mWidth + 2 Then Exit Function
    If mListFlipped Then
        If pointsY < mTop - DropListDepth() - 2 Then Exit Function
        If pointsY > mTop + mHeight + 2 Then Exit Function
    Else
        If pointsY < mTop - 2 Then Exit Function
        If pointsY > mTop + mHeight + DropListDepth() + 2 Then Exit Function
    End If
    ListAreaHolds = True
End Function

' True when this component projects floating chrome over the given
' sheet point: a visible modal overlay claims everything it covers, an
' open drop list claims its face plus the dropped rows and pagers.
Friend Function ChromeClaimsPoint( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double _
) As Boolean
    If Not mVisible Then Exit Function
    If mKind = KIND_OVERLAY Then
        If pointsX >= mLeft And pointsX <= mLeft + mWidth _
            And pointsY >= mTop And pointsY <= mTop + mHeight Then
            ChromeClaimsPoint = True
        End If
        Exit Function
    End If
    If (mKind = KIND_COMBO Or mKind = KIND_SELECT) And mSelectOpen Then
        ChromeClaimsPoint = ListAreaHolds(pointsX, pointsY)
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
    Dim pointsX As Double
    Dim pointsY As Double

    If Not PointerPoints(pointsX, pointsY) Then Exit Function
    If pointsX < mLeft - 8 Or pointsX > mLeft + mWidth + 8 Then Exit Function
    If pointsY < mTop - 4 Or pointsY > mTop + mHeight + 4 Then Exit Function
    CursorOverTrack = True
End Function

' Horizontal only: a drag keeps tracking while the cursor strays above or
' below the grid, where its vertical position cannot be mapped.
Private Function CursorTrackFraction() As Double
    Dim screenPoint As RdxScreenPoint
    Dim cursorPoints As Double
    Dim overriddenY As Double

    CursorTrackFraction = -1
    If Not ReDimUI.PointerOverride(cursorPoints, overriddenY) Then
        If GetCursorPos(screenPoint) = 0 Then Exit Function
        cursorPoints = ScreenXToSheetPoints(screenPoint.pixelX)
        If cursorPoints = UNMAPPABLE_POINTS Then Exit Function
    End If
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
    ApplyListCaret hostSheet, themeValue, 12
    SyncDropList hostSheet, themeValue
End Sub

' The open drop list of a SelectBox or ComboBox: at most mListRows item
' rows from mListOffset, with a clickable pager row at each edge that has
' items beyond it. Rows map back to item indexes through mListMatches.
' The combo's keyboard highlight paints in the accent, the current item
' in the muted fill. Each window row keeps the state it was last drawn
' with, so an arrow key that moves the highlight rewrites two rows, not
' the window; a row whose frame held rewrites only its fill, text, and
' ink, and a row is brought to the front only when first drawn.
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
    Dim rowTop As Double
    Dim rowFill As Long
    Dim rowInk As Long
    Dim rowKey As String
    Dim rowFrame As String
    Dim rowText As String
    Dim rowLook As String
    Dim boldAt As Long
    Dim boldLength As Long
    Dim priorKey As String
    Dim partCreated As Boolean
    Dim listSig As String
    Dim showEmpty As Boolean
    Dim listTopAt As Double

    ' A closed list draws nothing, so it skips the match walk: an item
    ' change on a closed long list would otherwise rebuild every match.
    ' Every path that opens the list sets its window and highlight.
    If mSelectOpen Then
        Set mListMatches = BuildListMatches(hostSheet)
        matchTotal = mListMatches.Count
        If mListOffset > matchTotal - mListRows Then
            mListOffset = matchTotal - mListRows
        End If
        If mListOffset < 0 Then mListOffset = 0
        If mListHighlight > matchTotal Then mListHighlight = 0
        If mVisible Then
            wanted = matchTotal - mListOffset
            If wanted > mListRows Then wanted = mListRows
            aboveCount = mListOffset
            belowCount = matchTotal - mListOffset - wanted
        End If
    End If
    If aboveCount > 0 Then rowShift = 1
    rowPitch = DropRowPitch()
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    showEmpty = (mSelectOpen And mVisible And matchTotal = 0)
    listTopAt = ListTop()
    ' Keystrokes and repaints that leave the visible list identical (same
    ' window, rows, highlight, selection, direction, geometry, and theme)
    ' skip the shape work entirely.
    listSig = CStr(ObjPtr(themeValue)) & ":" & CStr(wanted) & ":" & _
        CStr(mValue) & ":" & CStr(mListOffset) & ":" & _
        CStr(mListHighlight) & ":" & CStr(aboveCount) & ":" & _
        CStr(belowCount) & ":" & CStr(showEmpty) & ":" & _
        CStr(listTopAt) & ":" & CStr(mLeft) & ":" & CStr(mTop) & ":" & _
        CStr(mWidth) & ":" & CStr(mHeight) & ":" & CStr(resolvedSize) & _
        Chr$(2) & mListMatchesFilter
    For idx = 1 To wanted
        itemIndex = CLng(mListMatches.Item(mListOffset + idx))
        listSig = listSig & "|" & CStr(itemIndex) & "=" & ItemAt(itemIndex)
    Next idx
    If mAppliedValid And listSig = mAppliedListSig Then Exit Sub
    mAppliedListSig = listSig
    If Not mAppliedValid Then
        ' The 0.17.0 inert indicator row is superseded by the pagers.
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "optm").Delete
        On Error GoTo 0
    End If
    SyncListPager hostSheet, themeValue, "optu", ChrW(9650), _
        aboveCount, listTopAt, resolvedSize, mAppliedPagerUpSig
    SyncEmptyRow hostSheet, themeValue, showEmpty, listTopAt, rowPitch, _
        resolvedSize
    EnsureRowStates mAppliedRowStates, wanted
    For idx = 1 To wanted
        matchIndex = mListOffset + idx
        itemIndex = CLng(mListMatches.Item(matchIndex))
        rowTop = listTopAt + (rowShift + idx - 1) * rowPitch
        If matchIndex = mListHighlight Then
            rowFill = themeValue.PrimaryColor
            rowInk = themeValue.OnPrimaryColor
        ElseIf itemIndex = CLng(mValue) Then
            rowFill = themeValue.MutedColor
            rowInk = themeValue.OnSurfaceColor
        Else
            rowFill = themeValue.SurfaceColor
            rowInk = themeValue.OnSurfaceColor
        End If
        rowFrame = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            CStr(rowPitch) & ":" & CStr(resolvedSize)
        rowText = GutterText(itemIndex = CLng(mValue), ItemAt(itemIndex))
        ' A filtering combo bolds what the text matched in each row. The
        ' gutter leads every row and is never bold, so a text write
        ' clears an older match.
        boldAt = 0
        boldLength = 0
        If LenB(mListMatchesFilter) > 0 Then
            boldAt = InStr(1, ItemAt(itemIndex), mListMatchesFilter, vbTextCompare)
            If boldAt > 0 Then
                boldAt = boldAt + Len(rowText) - Len(ItemAt(itemIndex))
                boldLength = Len(mListMatchesFilter)
            End If
        End If
        ' The key ends with the bold span, after its own separator, so a
        ' row that kept its look can tell a moved match from a new row.
        rowLook = rowFrame & Chr$(1) & CStr(rowFill) & ":" & _
            CStr(rowInk) & ":" & rowText
        rowKey = rowLook & Chr$(2) & CStr(boldAt) & "+" & CStr(boldLength)
        priorKey = mAppliedRowStates(idx)
        If rowKey <> priorKey Then
            mAppliedRowStates(idx) = rowKey
            Set optionShape = EnsurePartShape( _
                hostSheet, "opt" & idx, msoShapeRectangle, partCreated)
            If Not partCreated And Left$(priorKey, Len(rowLook) + 1) _
                = rowLook & Chr$(2) Then
                ' Same row, the match moved: only its letters change.
                MoveRowBold optionShape, Mid$(priorKey, Len(rowLook) + 2), _
                    boldAt, boldLength
            ElseIf KeyFrame(priorKey) = rowFrame And Not partCreated Then
                ' Same frame: only the fill, the text, and the ink move.
                optionShape.Fill.ForeColor.RGB = rowFill
                With optionShape.TextFrame2.TextRange
                    .Text = rowText
                    .Font.Fill.ForeColor.RGB = rowInk
                    If boldAt > 0 Then
                        .Characters(boldAt, boldLength).Font.Bold = msoTrue
                    End If
                End With
            Else
                optionShape.Left = mLeft
                optionShape.Top = rowTop
                optionShape.Width = mWidth
                optionShape.Height = rowPitch
                optionShape.Fill.Visible = msoTrue
                optionShape.Fill.ForeColor.RGB = rowFill
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
                    .TextRange.Text = rowText
                    .TextRange.ParagraphFormat.Alignment = msoAlignLeft
                    With .TextRange.Font
                        .Name = themeValue.FontName
                        .Size = resolvedSize
                        .Bold = msoFalse
                        .Fill.ForeColor.RGB = rowInk
                    End With
                    If boldAt > 0 Then
                        .TextRange.Characters(boldAt, boldLength).Font.Bold = msoTrue
                    End If
                End With
                If LenB(priorKey) = 0 Or partCreated Then
                    EnsureGutterTab optionShape
                End If
                optionShape.Visible = msoTrue
                optionShape.OnAction = mAppliedOnAction
                If LenB(priorKey) = 0 Then optionShape.ZOrder msoBringToFront
            End If
        End If
    Next idx
    EnsureRowStates mAppliedRowStates, mAppliedOptionCount
    For idx = wanted + 1 To mAppliedOptionCount
        mAppliedRowStates(idx) = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & idx).Delete
        On Error GoTo 0
    Next idx
    mAppliedOptionCount = wanted
    SyncListPager hostSheet, themeValue, "optd", ChrW(9660), _
        belowCount, listTopAt + (rowShift + wanted) * rowPitch, _
        resolvedSize, mAppliedPagerDownSig
End Sub

' A drop row that kept its item while its match changed: the letters
' that left the match lose bold and the ones that joined it gain it.
' priorSpan is the span the row was drawn with, as "at+length".
Private Sub MoveRowBold( _
    ByVal rowShape As Shape, _
    ByVal priorSpan As String, _
    ByVal boldAt As Long, _
    ByVal boldLength As Long _
)
    Dim priorAt As Long
    Dim priorLength As Long
    Dim splitAt As Long

    splitAt = InStr(1, priorSpan, "+")
    priorAt = CLng(Left$(priorSpan, splitAt - 1))
    priorLength = CLng(Mid$(priorSpan, splitAt + 1))
    With rowShape.TextFrame2.TextRange
        If priorAt = boldAt And boldAt > 0 Then
            If boldLength > priorLength Then
                .Characters(boldAt + priorLength, boldLength - priorLength) _
                    .Font.Bold = msoTrue
            ElseIf boldLength < priorLength Then
                .Characters(boldAt + boldLength, priorLength - boldLength) _
                    .Font.Bold = msoFalse
            End If
        Else
            If priorAt > 0 Then .Characters(priorAt, priorLength).Font.Bold = msoFalse
            If boldAt > 0 Then .Characters(boldAt, boldLength).Font.Bold = msoTrue
        End If
    End With
End Sub

' A list row's text: the check in the gutter when the row is marked,
' then a tab and the item, so marked and unmarked rows align.
Private Function GutterText(ByVal marked As Boolean, ByVal itemText As String) As String
    If marked Then
        GutterText = ChrW(10003) & vbTab & itemText
    Else
        GutterText = vbTab & itemText
    End If
End Function

' The tab stop the gutter aligns to. Shape text keeps it through later
' text writes, so rows check for it only when first drawn.
Private Sub EnsureGutterTab(ByVal rowShape As Shape)
    With rowShape.TextFrame2.TextRange.ParagraphFormat.TabStops
        If .Count = 0 Then .Add msoTabStopLeft, LIST_GUTTER
    End With
End Sub

' The inert row an open list shows when it has nothing to list, in
' muted italic: "No matches" under a combo's filter, "No items" in
' an empty select.
Private Sub SyncEmptyRow( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal showIt As Boolean, _
    ByVal rowTop As Double, _
    ByVal rowPitch As Double, _
    ByVal resolvedSize As Double _
)
    Dim emptyShape As Shape
    Dim emptyKey As String
    Dim emptyText As String

    If showIt Then
        If mKind = KIND_COMBO Then
            emptyText = "No matches"
        Else
            emptyText = "No items"
        End If
        emptyKey = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            CStr(rowPitch) & ":" & CStr(resolvedSize) & ":" & emptyText
        If emptyKey = mAppliedEmptyRowSig Then Exit Sub
        mAppliedEmptyRowSig = emptyKey
        Set emptyShape = EnsurePartShape(hostSheet, "optn", msoShapeRectangle)
        emptyShape.Left = mLeft
        emptyShape.Top = rowTop
        emptyShape.Width = mWidth
        emptyShape.Height = rowPitch
        emptyShape.Fill.Visible = msoTrue
        emptyShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        emptyShape.Line.Visible = msoTrue
        emptyShape.Line.ForeColor.RGB = themeValue.BorderColor
        emptyShape.Line.Weight = 1
        With emptyShape.TextFrame2
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 2
            .MarginBottom = 2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .TextRange.Text = GutterText(False, emptyText)
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            With .TextRange.Font
                .Name = themeValue.FontName
                .Size = resolvedSize
                .Bold = msoFalse
                .Italic = msoTrue
                .Fill.ForeColor.RGB = themeValue.OnMutedColor
            End With
        End With
        EnsureGutterTab emptyShape
        emptyShape.Visible = msoTrue
        emptyShape.OnAction = mAppliedOnAction
        emptyShape.ZOrder msoBringToFront
    ElseIf LenB(mAppliedEmptyRowSig) > 0 Then
        mAppliedEmptyRowSig = vbNullString
        On Error Resume Next
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "optn").Delete
        On Error GoTo 0
    End If
End Sub

' The top edge of the open list: right under the face, or above it with
' its last row touching the face when the list opened upward.
Private Function ListTop() As Double
    If mListFlipped Then
        ListTop = mTop - DropListDepth()
    Else
        ListTop = mTop + mHeight
    End If
End Function

' A list that would run past the bottom of the visible window opens
' upward instead, when the window has more room above the face than
' below it. Decided once per opening, so a filter that shrinks the list
' never makes it jump sides.
Private Function ListOpensUpward() As Boolean
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim shownRows As Long
    Dim fullDepth As Double
    Dim roomBelow As Double

    shownRows = ItemCount
    If shownRows > mListRows Then shownRows = mListRows
    If shownRows < 1 Then shownRows = 1
    fullDepth = shownRows * DropRowPitch()
    If ItemCount > mListRows Then fullDepth = fullDepth + 2 * mHeight
    OwnerApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    roomBelow = viewTop + viewHeight - (mTop + mHeight)
    If fullDepth <= roomBelow Then Exit Function
    ListOpensUpward = (mTop - viewTop > roomBelow)
End Function

' The frame half of a stored row key (frame & Chr$(1) & content): when a
' row's new frame matches, only its content needs writing.
Private Function KeyFrame(ByVal storedKey As String) As String
    Dim splitAt As Long

    splitAt = InStr(1, storedKey, Chr$(1))
    If splitAt > 0 Then KeyFrame = Left$(storedKey, splitAt - 1)
End Function

' Grows a row-state array to hold at least rowsNeeded rows, keeping what
' it holds; an erased or never-sized array starts empty.
Private Sub EnsureRowStates(ByRef rowStates() As String, ByVal rowsNeeded As Long)
    Dim rowsHeld As Long

    If rowsNeeded < 1 Then Exit Sub
    On Error Resume Next
    rowsHeld = UBound(rowStates)
    Err.Clear
    On Error GoTo 0
    If rowsHeld >= rowsNeeded Then Exit Sub
    If rowsHeld = 0 Then
        ReDim rowStates(1 To rowsNeeded)
    Else
        ReDim Preserve rowStates(1 To rowsNeeded)
    End If
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
    ' An empty list still shows its one inert row.
    If matchTotal = 0 Then
        DropListDepth = DropRowPitch()
        Exit Function
    End If
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
    ' Matches are cached per filter text; any item change drops them.
    If Not mListMatches Is Nothing Then
        If filterText = mListMatchesFilter Then
            Set BuildListMatches = mListMatches
            Exit Function
        End If
    End If
    mListMatchesFilter = filterText
    Set BuildListMatches = FilterItems(filterText)
End Function

' Opens the list pristine, scrolled so the current item (the SelectBox
' selection, or the item a combo's text names) is the first row; the
' render clamps the window near the end of the list.
Private Sub OpenDropList(ByVal hostSheet As Worksheet)
    Dim currentItem As Long

    mSelectOpen = True
    NoteListOpened
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
' lies beyond. appliedKey holds the state it was last drawn with: empty
' when absent, "?" when unknown after Render, which forces one redraw or
' one delete.
Private Sub SyncListPager( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal arrowGlyph As String, _
    ByVal beyondCount As Long, _
    ByVal rowTop As Double, _
    ByVal resolvedSize As Double, _
    ByRef appliedKey As String _
)
    Dim pagerShape As Shape
    Dim pagerKey As String
    Dim firstDraw As Boolean

    If beyondCount > 0 Then
        pagerKey = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            CStr(beyondCount) & ":" & CStr(resolvedSize)
        If pagerKey = appliedKey Then Exit Sub
        firstDraw = (LenB(appliedKey) = 0 Or appliedKey = "?")
        appliedKey = pagerKey
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
        If firstDraw Then pagerShape.ZOrder msoBringToFront
    ElseIf LenB(appliedKey) > 0 Then
        appliedKey = vbNullString
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

' Kinds that take keyboard focus: every interactive control, a text
' field only when it floats (a cell-backed one lives in the grid), and
' an image only when it has a click handler.
Friend Function IsFocusable() As Boolean
    If Not mVisible Or Not mEnabled Or mBusy Then Exit Function
    Select Case mKind
        Case KIND_BUTTON, KIND_TOGGLE, KIND_TICKBOX, KIND_RADIO, _
            KIND_STEPPER, KIND_SLIDEBAR, KIND_SELECT, KIND_CHECKLIST, _
            KIND_TRANSFER
            IsFocusable = True
        Case KIND_INPUT, KIND_COMBO
            IsFocusable = IsFloatField()
        Case KIND_IMAGE
            IsFocusable = (LenB(mClickProc) > 0)
    End Select
End Function

Friend Function IsTabStop() As Boolean
    IsTabStop = IsFocusable() And mTabIndex >= 0
End Function

Friend Property Get TabPosition() As Long
    TabPosition = mTabIndex
End Property

' Takes keyboard focus. A float field focuses as an editor, opening a
' combo's list only when openList says so (a click does, Tab does
' not); any other control takes the keys and shows a focus ring.
Friend Sub TakeFocus(ByVal hostApp As ReDimUI, ByVal openList As Boolean)
    If ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    If IsFloatField() Then
        FocusField hostApp, openList
        Exit Sub
    End If
    mFieldSelAtFocus = CurrentSelectionKey()
    mTypeAhead = vbNullString
    If Not mKeyCursorSet Then
        mKeyRow = 1
        mKeyPanel = 0
        mKeyCursorSet = True
    End If
    ReDimUI.SetKeyboardFocus hostApp, Me
    Touch
End Sub

' Gives keyboard focus up: a float field blurs, committing or
' reverting; any other control drops its ring and closes an open list.
Friend Sub LeaveFocus(ByVal hostApp As ReDimUI, ByVal keepEdits As Boolean)
    If IsFloatField() Then
        BlurField hostApp, keepEdits
        Exit Sub
    End If
    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    ReDimUI.ClearKeyboardFocus
    mSelectOpen = False
    mListHighlight = 0
    Touch
End Sub

' Keyboard focus that lands outside the visible window scrolls it into
' view, by rows and columns so the selection stays where it is.
Friend Sub ScrollIntoSight()
    Dim hostApp As ReDimUI
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim mainShape As Shape
    Dim scrollTo As Long

    Set hostApp = OwnerApp
    On Error GoTo Done
    If Not hostApp.Sheet Is ActiveSheet Then Exit Sub
    hostApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    Set mainShape = hostApp.Sheet.Shapes(ShapeName)
    ' Above or left of the window: align to the top or left edge. Below
    ' or right: scroll just far enough to show the far edge.
    If mTop < viewTop Then
        ActiveWindow.ScrollRow = mainShape.TopLeftCell.Row
    ElseIf mTop + mHeight > viewTop + viewHeight * 0.95 Then
        scrollTo = ActiveWindow.ScrollRow
        Do While hostApp.Sheet.Rows(scrollTo).Top + viewHeight * 0.95 _
            < mTop + mHeight And scrollTo < mainShape.TopLeftCell.Row
            scrollTo = scrollTo + 1
        Loop
        ActiveWindow.ScrollRow = scrollTo
    End If
    If mLeft < viewLeft Then
        ActiveWindow.ScrollColumn = mainShape.TopLeftCell.Column
    ElseIf mLeft + mWidth > viewLeft + viewWidth * 0.95 Then
        scrollTo = ActiveWindow.ScrollColumn
        Do While hostApp.Sheet.Columns(scrollTo).Left + viewWidth * 0.95 _
            < mLeft + mWidth And scrollTo < mainShape.TopLeftCell.Column
            scrollTo = scrollTo + 1
        Loop
        ActiveWindow.ScrollColumn = scrollTo
    End If
Done:
End Sub

Friend Sub FocusField(ByVal hostApp As ReDimUI, Optional ByVal openList As Boolean = True)
    If Not IsFloatField() Then Exit Sub
    mFieldTextAtFocus = mText
    mFieldSelAtFocus = CurrentSelectionKey()
    mCaretPos = Len(mText)
    mSelAnchor = mCaretPos
    Set mUndoStack = Nothing
    Set mRedoStack = Nothing
    mLastEditKind = vbNullString
    mCaretOn = True
    mCaretAccumMs = 0
    ReDimUI.SetKeyboardFocus hostApp, Me
    If mKind = KIND_COMBO And openList Then OpenDropList hostApp.Sheet
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
    mSelAnchor = mCaretPos
    ' A pending OnInput reports the final text before the commit.
    If mInputDueMs > 0 Then FireInput hostApp
    If keepEdits Then
        If mKind = KIND_COMBO Then
            exactAt = ExactItemMatch(mText)
            If exactAt = 0 And mRestrictToItems And LenB(mText) > 0 Then
                ' A restricted combo takes the first item the text
                ' begins, or goes back to its text at focus.
                exactAt = PrefixItemMatch(mText)
                If exactAt = 0 Then
                    mText = mFieldTextAtFocus
                    exactAt = ExactItemMatch(mText)
                End If
            End If
            If exactAt > 0 Then
                mValue = exactAt
                mText = ItemAt(exactAt)
            Else
                mValue = 0
            End If
        End If
        RunValidation hostApp
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
    Dim extendSelection As Boolean
    Dim pastedText As String

    If Not IsFloatField() Then
        HandleControlKey hostApp, keyText
        Exit Sub
    End If
    If mCaretPos > Len(mText) Then mCaretPos = Len(mText)
    If mSelAnchor < 0 Or mSelAnchor > Len(mText) Then mSelAnchor = mCaretPos
    textBeforeKey = mText
    ' Shift-extended keys arrive as {SHIFTLEFT} and the like: the plain
    ' key's move, with the selection's anchor kept.
    If Left$(keyText, 6) = "{SHIFT" Then
        extendSelection = True
        keyText = "{" & Mid$(keyText, 7)
    End If
    Select Case keyText
        Case "{ENTER}"
            ' Single-line: Enter commits - except a combo with a
            ' highlighted option, which takes that option. Multi-line:
            ' Enter is a newline and committing moves to Tab,
            ' Ctrl+Enter, or clicking away.
            If mMultiLine Then
                RecordUndo "type"
                ReplaceSelection vbLf
            ElseIf mKind = KIND_COMBO And mSelectOpen _
                And mListHighlight >= 1 Then
                ' A keyboard pick commits exactly like a mouse pick,
                ' handlers included, and only when the value changes.
                If PickComboItem(hostApp, mListHighlight) Then
                    RunClickHandlers hostApp
                End If
                Exit Sub
            Else
                ' Enter commits, then clicks the app's default button.
                BlurField hostApp, True
                hostApp.ClickDefaultButton
                Exit Sub
            End If
        Case "{TAB}", "{BACKTAB}"
            ' Tab takes a combo's suggestion, commits, and moves focus
            ' on; with no other control to move to, it commits and
            ' leaves.
            If keyText = "{TAB}" Then
                If AcceptGhost(hostApp.Sheet) Then NoteFieldEdit hostApp
            End If
            If Not hostApp.MoveFocusFrom(mComponentId, keyText = "{TAB}") Then
                BlurField hostApp, True
            End If
            Exit Sub
        Case "{CTRLENTER}"
            BlurField hostApp, True
            Exit Sub
        Case "{ESC}"
            ' Esc closes an open list first, keeping focus and text;
            ' with the list closed it reverts the field to the text it
            ' had when focus arrived and leaves, firing nothing.
            If mKind = KIND_COMBO And mSelectOpen Then
                CloseDropList
            Else
                BlurField hostApp, False
            End If
            Exit Sub
        Case "{ALTUP}"
            If mKind = KIND_COMBO And mSelectOpen Then CloseDropList
            Exit Sub
        Case "{ALTDOWN}", "{F4}"
            If mKind = KIND_COMBO And Not mSelectOpen Then
                OpenDropList hostApp.Sheet
            End If
        Case "{PGUP}", "{PGDN}"
            If mKind = KIND_COMBO Then
                ComboListNav hostApp, IIf(keyText = "{PGDN}", mListRows, -mListRows)
            End If
        Case "{BS}", "{DEL}", "{WORDBS}", "{WORDDEL}"
            ' Deletes take the selection when there is one; otherwise
            ' a character, or a word with Ctrl, before or after.
            If Not HasSelection() Then
                Select Case keyText
                    Case "{BS}"
                        mSelAnchor = mCaretPos - 1
                    Case "{DEL}"
                        mSelAnchor = mCaretPos + 1
                    Case "{WORDBS}"
                        mSelAnchor = WordStartBefore(mCaretPos)
                    Case Else
                        mSelAnchor = WordStartAfter(mCaretPos)
                End Select
                If mSelAnchor < 0 Then mSelAnchor = 0
                If mSelAnchor > Len(mText) Then mSelAnchor = Len(mText)
            End If
            If HasSelection() Then
                RecordUndo "delete"
                ReplaceSelection vbNullString
            End If
        Case "{LEFT}"
            If HasSelection() And Not extendSelection Then
                MoveCaretTo SelStart(), False
            Else
                MoveCaretTo mCaretPos - 1, extendSelection
            End If
        Case "{RIGHT}"
            ' Right at the end of a combo's text takes its suggestion.
            If HasSelection() And Not extendSelection Then
                MoveCaretTo SelStart() + SelLength(), False
            ElseIf extendSelection Then
                MoveCaretTo mCaretPos + 1, True
            ElseIf Not AcceptGhost(hostApp.Sheet) Then
                MoveCaretTo mCaretPos + 1, False
            End If
        Case "{WORDLEFT}"
            MoveCaretTo WordStartBefore(mCaretPos), extendSelection
        Case "{WORDRIGHT}"
            MoveCaretTo WordStartAfter(mCaretPos), extendSelection
        Case "{UP}", "{DOWN}"
            If mMultiLine Then
                MoveCaretLine IIf(keyText = "{UP}", -1, 1)
                MoveCaretTo mCaretPos, extendSelection
            ElseIf mKind = KIND_COMBO And Not extendSelection Then
                ComboListNav hostApp, IIf(keyText = "{UP}", -1, 1)
            End If
        Case "{HOME}", "{END}"
            MoveCaretLineEdge (keyText = "{END}")
            MoveCaretTo mCaretPos, extendSelection
        Case "{TEXTHOME}"
            MoveCaretTo 0, extendSelection
        Case "{TEXTEND}"
            MoveCaretTo Len(mText), extendSelection
        Case "{SELECTALL}"
            mSelAnchor = 0
            mCaretPos = Len(mText)
            mLastEditKind = vbNullString
        Case "{COPY}"
            If HasSelection() Then
                WriteClipboardText Mid$(mText, SelStart() + 1, SelLength())
            End If
        Case "{CUT}"
            If HasSelection() Then
                If WriteClipboardText(Mid$(mText, SelStart() + 1, SelLength())) Then
                    RecordUndo "cut"
                    ReplaceSelection vbNullString
                End If
            End If
        Case "{PASTE}"
            pastedText = ReadClipboardText()
            If LenB(pastedText) > 0 Then
                RecordUndo "paste"
                ReplaceSelection pastedText
            End If
        Case "{UNDO}"
            StepHistory True
        Case "{REDO}"
            StepHistory False
        Case "{APOS}"
            RecordUndo "type"
            ReplaceSelection "'"
        Case "{QUOTE}"
            RecordUndo "type"
            ReplaceSelection """"
        Case Else
            ' Named keys a field has no use for are dropped, not typed.
            If Len(keyText) > 1 And Left$(keyText, 1) = "{" Then Exit Sub
            RecordUndo "type"
            ReplaceSelection keyText
    End Select
    mCaretOn = True
    mCaretAccumMs = 0
    If mText <> textBeforeKey Then NoteFieldEdit hostApp
    ' Live filtering, the payoff cells could never give: an edit leaves
    ' the pristine full list and re-filters by the new text, rewinding
    ' the window and dropping the keyboard highlight. Keys that change
    ' nothing reopen a closed list as it was.
    If mKind = KIND_COMBO Then
        If mText <> textBeforeKey Then
            If Not mSelectOpen Then NoteListOpened
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

Private Function HasSelection() As Boolean
    HasSelection = (mSelAnchor <> mCaretPos)
End Function

Private Function SelStart() As Long
    If mSelAnchor < mCaretPos Then
        SelStart = mSelAnchor
    Else
        SelStart = mCaretPos
    End If
End Function

Private Function SelLength() As Long
    SelLength = Abs(mCaretPos - mSelAnchor)
End Function

' Replaces the selection, or inserts at the caret when there is none,
' with what the field admits of the given text; the caret lands after
' it with nothing selected.
Private Sub ReplaceSelection(ByVal insertText As String)
    Dim keptBefore As String
    Dim keptAfter As String

    keptBefore = Left$(mText, SelStart())
    keptAfter = Mid$(mText, SelStart() + SelLength() + 1)
    insertText = FieldAdmits(insertText, keptBefore, keptAfter)
    mText = keptBefore & insertText & keptAfter
    mCaretPos = Len(keptBefore) + Len(insertText)
    mSelAnchor = mCaretPos
End Sub

' What a field takes of typed or pasted text, given the text around
' the insertion: CR and CRLF become line breaks, a trailing break (a
' copied cell brings one) is dropped, a single-line field turns breaks
' and tabs into spaces, Numeric keeps what a number can hold, and
' MaxLength cuts what does not fit.
Private Function FieldAdmits( _
    ByVal givenText As String, _
    ByVal keptBefore As String, _
    ByVal keptAfter As String _
) As String
    Dim room As Long
    givenText = Replace(givenText, vbCrLf, vbLf)
    givenText = Replace(givenText, vbCr, vbLf)
    If Len(givenText) > 1 And Right$(givenText, 1) = vbLf Then
        givenText = Left$(givenText, Len(givenText) - 1)
    End If
    If Not mMultiLine Then
        givenText = Replace(givenText, vbLf, " ")
        givenText = Replace(givenText, vbTab, " ")
    End If
    If mNumeric Then givenText = NumericAdmits(givenText, keptBefore, keptAfter)
    If mMaxLength > 0 Then
        room = mMaxLength - Len(keptBefore) - Len(keptAfter)
        If room < 0 Then room = 0
        If Len(givenText) > room Then givenText = Left$(givenText, room)
    End If
    FieldAdmits = givenText
End Function

' Numeric's filter: digits, one decimal separator (either . or , types
' the locale's), and a minus only at the very start.
Private Function NumericAdmits( _
    ByVal givenText As String, _
    ByVal keptBefore As String, _
    ByVal keptAfter As String _
) As String
    Dim separatorChar As String
    Dim hasSeparator As Boolean
    Dim idx As Long
    Dim oneChar As String

    If Application.UseSystemSeparators Then
        separatorChar = Application.International(xlDecimalSeparator)
    Else
        separatorChar = Application.DecimalSeparator
    End If
    hasSeparator = InStr(keptBefore & keptAfter, separatorChar) > 0
    For idx = 1 To Len(givenText)
        oneChar = Mid$(givenText, idx, 1)
        If oneChar Like "#" Then
            NumericAdmits = NumericAdmits & oneChar
        ElseIf mNumericDecimal And Not hasSeparator _
            And (oneChar = "." Or oneChar = "," Or oneChar = separatorChar) Then
            NumericAdmits = NumericAdmits & separatorChar
            hasSeparator = True
        ElseIf mNumericNegative And oneChar = "-" And LenB(keptBefore) = 0 _
            And LenB(NumericAdmits) = 0 And Left$(keptAfter, 1) <> "-" Then
            NumericAdmits = "-"
        End If
    Next idx
End Function

' Moves the caret. Extending keeps the selection's anchor; otherwise
' the selection collapses onto the caret. Either ends a typing run.
Private Sub MoveCaretTo(ByVal newPos As Long, ByVal extendSelection As Boolean)
    If newPos < 0 Then newPos = 0
    If newPos > Len(mText) Then newPos = Len(mText)
    mCaretPos = newPos
    If Not extendSelection Then mSelAnchor = newPos
    mLastEditKind = vbNullString
End Sub

' Undo steps: a run of typing is one step until a second's pause or any
' other kind of edit; each delete, cut, or paste is a step of its own.
' The history holds a hundred steps and a new edit clears the redo.
Private Sub RecordUndo(ByVal editKind As String)
    Dim nowMs As LongLong

    nowMs = GetTickCount64()
    If editKind <> "type" Or mLastEditKind <> "type" _
        Or nowMs - mLastEditMs > 1000 Then
        If mUndoStack Is Nothing Then Set mUndoStack = New Collection
        mUndoStack.Add Array(mText, mCaretPos, mSelAnchor)
        If mUndoStack.Count > 100 Then mUndoStack.Remove 1
        Set mRedoStack = New Collection
    End If
    mLastEditKind = editKind
    mLastEditMs = nowMs
End Sub

' Undo (backward) or redo: the current state goes on the other stack.
Private Sub StepHistory(ByVal undoing As Boolean)
    Dim fromStack As Collection
    Dim toStack As Collection
    Dim snapshot As Variant

    If mUndoStack Is Nothing Then Set mUndoStack = New Collection
    If mRedoStack Is Nothing Then Set mRedoStack = New Collection
    If undoing Then
        Set fromStack = mUndoStack
        Set toStack = mRedoStack
    Else
        Set fromStack = mRedoStack
        Set toStack = mUndoStack
    End If
    If fromStack.Count = 0 Then Exit Sub
    toStack.Add Array(mText, mCaretPos, mSelAnchor)
    snapshot = fromStack.Item(fromStack.Count)
    fromStack.Remove fromStack.Count
    mText = snapshot(0)
    mCaretPos = snapshot(1)
    mSelAnchor = snapshot(2)
    mLastEditKind = vbNullString
End Sub

Private Function IsWordChar(ByVal oneChar As String) As Boolean
    If LenB(oneChar) = 0 Then Exit Function
    IsWordChar = (oneChar Like "[A-Za-z0-9_]") Or AscW(oneChar) > 127
End Function

' Ctrl+Left: back over spaces and punctuation, then over the word
' before them, to where that word starts.
Private Function WordStartBefore(ByVal fromPos As Long) As Long
    Dim pos As Long

    pos = fromPos
    Do While pos > 0
        If IsWordChar(Mid$(mText, pos, 1)) Then Exit Do
        pos = pos - 1
    Loop
    Do While pos > 0
        If Not IsWordChar(Mid$(mText, pos, 1)) Then Exit Do
        pos = pos - 1
    Loop
    WordStartBefore = pos
End Function

' Ctrl+Right: over the rest of the word, then over the spaces after
' it, to where the next word starts, as Windows moves.
Private Function WordStartAfter(ByVal fromPos As Long) As Long
    Dim pos As Long

    pos = fromPos
    Do While pos < Len(mText)
        If Not IsWordChar(Mid$(mText, pos + 1, 1)) Then Exit Do
        pos = pos + 1
    Loop
    Do While pos < Len(mText)
        If IsWordChar(Mid$(mText, pos + 1, 1)) Then Exit Do
        pos = pos + 1
    Loop
    WordStartAfter = pos
End Function

' Unicode text on the clipboard, or "" when it holds none or is busy.
Private Function ReadClipboardText() As String
    Dim memoryHandle As LongPtr
    Dim textAddress As LongPtr
    Dim charCount As Long

    If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
    If OpenClipboard(0) = 0 Then Exit Function
    memoryHandle = GetClipboardData(CF_UNICODETEXT)
    If memoryHandle <> 0 Then
        textAddress = GlobalLock(memoryHandle)
        If textAddress <> 0 Then
            charCount = lstrlenW(textAddress)
            If charCount > 0 Then
                ReadClipboardText = String$(charCount, vbNullChar)
                CopyMemory StrPtr(ReadClipboardText), textAddress, charCount * 2
            End If
            GlobalUnlock memoryHandle
        End If
    End If
    CloseClipboard
End Function

' Puts Unicode text on the clipboard. True when it landed.
Private Function WriteClipboardText(ByVal copiedText As String) As Boolean
    Dim memoryHandle As LongPtr
    Dim textAddress As LongPtr
    Dim byteCount As LongPtr

    If OpenClipboard(0) = 0 Then Exit Function
    EmptyClipboard
    byteCount = (Len(copiedText) + 1) * 2
    memoryHandle = GlobalAlloc(GMEM_MOVEABLE_ZEROED, byteCount)
    If memoryHandle <> 0 Then
        textAddress = GlobalLock(memoryHandle)
        If textAddress <> 0 Then
            If Len(copiedText) > 0 Then
                CopyMemory textAddress, StrPtr(copiedText), byteCount - 2
            End If
            GlobalUnlock memoryHandle
            If SetClipboardData(CF_UNICODETEXT, memoryHandle) <> 0 Then
                WriteClipboardText = True
            Else
                GlobalFree memoryHandle
            End If
        Else
            GlobalFree memoryHandle
        End If
    End If
    CloseClipboard
End Function

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

' Outside-press blur for whichever control holds keyboard focus,
' watched by the same pump frames as slider drags.
' Two independent signals, because neither alone covers everything:
' - The selection moved since focus: a cell was clicked or arrowed to.
'   The press itself can be invisible to polling (the grid's selection
'   mouse loop holds WM_TIMER until the button is back up), but the
'   selection it leaves behind is durable state the next frame sees.
'   Event-free, so it works even when a host disables application
'   events, and it runs in proven tick context.
' - A press edge lands off the control: shapes, chrome, other windows -
'   clicks that never touch the selection.
Friend Sub TickFocusWatch( _
    ByVal hostApp As ReDimUI, _
    ByVal pressEdge As Boolean, _
    ByVal pollSelection As Boolean _
)
    If Not ReDimUI.IsComponentFocused(hostApp.AppId, mComponentId) Then Exit Sub
    ' A modal's buttons keep focus until the dialog closes.
    If Left$(mComponentId, 4) = "mdl_" Then Exit Sub
    ' The selection compare costs three COM reads, so it rides the work
    ' cadence; a 50 ms commit latency is imperceptible. The press edge
    ' stays at frame rate.
    If pollSelection Then
        If CurrentSelectionKey() <> mFieldSelAtFocus Then
            LeaveFocus hostApp, True
            Exit Sub
        End If
    End If
    If Not pressEdge Then Exit Sub
    If CursorOverField() Then Exit Sub
    LeaveFocus hostApp, True
End Sub

' An open list that no field watch covers (a SelectBox, or a cell-backed
' combo's suggestions) closes the way a native drop-down does: on a press
' off its face and rows, or when the grid selection moves. The first frame
' records the selection, since a cell combo opens from an edit whose Enter
' moves the selection right after.
Friend Function NeedsListWatch() As Boolean
    If Not mSelectOpen Or Not mVisible Then Exit Function
    NeedsListWatch = (mKind = KIND_SELECT) _
        Or (mKind = KIND_COMBO And Not IsFloatField())
End Function

Friend Sub TickListWatch(ByVal pressEdge As Boolean, ByVal pollSelection As Boolean)
    If Not NeedsListWatch() Then Exit Sub
    If pollSelection Then
        If LenB(mListSelAtOpen) = 0 Then
            mListSelAtOpen = CurrentSelectionKey()
        ElseIf CurrentSelectionKey() <> mListSelAtOpen Then
            CloseDropList
            Exit Sub
        End If
    End If
    If pressEdge Then
        If Not CursorOverField() Then CloseDropList
    End If
End Sub

' Closes an open list without picking anything.
Friend Sub CloseDropList()
    If Not mSelectOpen Then Exit Sub
    mSelectOpen = False
    mListHighlight = 0
    Touch
End Sub

' Every path that opens the list: the app's other lists close, the list
' picks its side, the watch starts over, and the pump runs so the watch
' sees presses elsewhere.
Private Sub NoteListOpened()
    mListSelAtOpen = vbNullString
    mListFlipped = ListOpensUpward()
    OwnerApp.CloseListsExcept mComponentId
    ReDimUI.RequestPump
End Sub

Private Function CursorOverField() As Boolean
    Dim pointsX As Double
    Dim pointsY As Double

    If Not PointerPoints(pointsX, pointsY) Then Exit Function
    ' The open option list counts as part of the field.
    CursorOverField = ListAreaHolds(pointsX, pointsY)
End Function

' Item indexes whose text contains the filter, case-insensitively; every
' index when the filter is empty.
Private Function FilterItems(ByVal filterText As String) As Collection
    Dim idx As Long

    Set FilterItems = New Collection
    If mItems Is Nothing Then Exit Function
    If mItems.Count = 0 Then Exit Function
    EnsureItemTexts
    If LenB(filterText) = 0 Then
        For idx = 1 To mItems.Count
            FilterItems.Add idx
        Next idx
    Else
        For idx = 1 To mItems.Count
            If InStr(1, mItemTexts(idx), filterText, vbTextCompare) > 0 Then
                FilterItems.Add idx
            End If
        Next idx
    End If
End Function

' The first item that begins with the text, in any case.
Private Function PrefixItemMatch(ByVal candidate As String) As Long
    Dim idx As Long

    If LenB(candidate) = 0 Then Exit Function
    If mItems Is Nothing Then Exit Function
    If mItems.Count = 0 Then Exit Function
    EnsureItemTexts
    For idx = 1 To mItems.Count
        If StrComp(Left$(mItemTexts(idx), Len(candidate)), candidate, _
            vbTextCompare) = 0 Then
            PrefixItemMatch = idx
            Exit Function
        End If
    Next idx
End Function

' A combo filtering its list suggests the first item that begins with
' the typed text, unless an item is that text already: the rest of it
' shows after the caret in muted ink, and Right at the end or Tab takes
' it.
Private Function GhostItem(ByVal hostSheet As Worksheet) As Long
    Dim matchEntry As Variant
    Dim itemText As String
    Dim typedLength As Long

    If mKind <> KIND_COMBO Then Exit Function
    If Not mSelectOpen Or mFilterPristine Then Exit Function
    typedLength = Len(mText)
    If typedLength = 0 Or mCaretPos <> typedLength Then Exit Function
    If HasSelection() Then Exit Function
    Set mListMatches = BuildListMatches(hostSheet)
    For Each matchEntry In mListMatches
        itemText = ItemAt(CLng(matchEntry))
        If Len(itemText) = typedLength Then
            If StrComp(itemText, mText, vbTextCompare) = 0 Then
                GhostItem = 0
                Exit Function
            End If
        ElseIf GhostItem = 0 And Len(itemText) > typedLength Then
            If StrComp(Left$(itemText, typedLength), mText, vbTextCompare) = 0 Then
                GhostItem = CLng(matchEntry)
            End If
        End If
    Next matchEntry
End Function

Private Function GhostRest(ByVal hostSheet As Worksheet) As String
    Dim ghostAt As Long

    ghostAt = GhostItem(hostSheet)
    If ghostAt > 0 Then GhostRest = Mid$(ItemAt(ghostAt), Len(mText) + 1)
End Function

' Takes the suggested item whole, in the item's own spelling.
Private Function AcceptGhost(ByVal hostSheet As Worksheet) As Boolean
    Dim ghostAt As Long

    ghostAt = GhostItem(hostSheet)
    If ghostAt = 0 Then Exit Function
    RecordUndo "complete"
    mSelAnchor = 0
    mCaretPos = Len(mText)
    ReplaceSelection ItemAt(ghostAt)
    AcceptGhost = True
End Function

Private Function ExactItemMatch(ByVal candidate As String) As Long
    Dim idx As Long

    If mItems Is Nothing Then Exit Function
    If mItems.Count = 0 Then Exit Function
    EnsureItemTexts
    For idx = 1 To mItems.Count
        If StrComp(mItemTexts(idx), candidate, vbTextCompare) = 0 Then
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
    ' Kinds drawn as a single shape delete it by name. Composite kinds
    ' delete the main shape and every part by prefix in one pass, where
    ' guessing each part name cost a failed lookup per guess.
    If HasParts() Then
        DeleteShapesNamed hostSheet, ShapeName & PART_SEPARATOR, ShapeName
    Else
        On Error Resume Next
        hostSheet.Shapes(ShapeName).Delete
        On Error GoTo 0
    End If
    mAppliedOptionCount = 0
    ForgetPartSignatures
End Sub

Private Function HasParts() As Boolean
    Select Case mKind
        Case KIND_LABEL, KIND_CARD, KIND_SPINNER, KIND_OVERLAY
            HasParts = False
        Case Else
            HasParts = True
    End Select
End Function

' Deletes every shape on the sheet whose name starts with namePrefix, or
' equals exactName when one is given, through one ShapeRange call; a
' shape that refuses falls back to a delete of its own.
Private Sub DeleteShapesNamed( _
    ByVal hostSheet As Worksheet, _
    ByVal namePrefix As String, _
    ByVal exactName As String _
)
    Dim candidateShape As Shape
    Dim doomedNames() As String
    Dim doomedCount As Long
    Dim idx As Long
    Dim shapeLabel As String

    If hostSheet.Shapes.Count = 0 Then Exit Sub
    ReDim doomedNames(1 To hostSheet.Shapes.Count)
    For Each candidateShape In hostSheet.Shapes
        shapeLabel = candidateShape.Name
        If Left$(shapeLabel, Len(namePrefix)) = namePrefix _
            Or (LenB(exactName) > 0 And shapeLabel = exactName) Then
            doomedCount = doomedCount + 1
            doomedNames(doomedCount) = shapeLabel
        End If
    Next candidateShape
    If doomedCount = 0 Then Exit Sub
    ReDim Preserve doomedNames(1 To doomedCount)
    On Error Resume Next
    Err.Clear
    hostSheet.Shapes.Range(doomedNames).Delete
    If Err.Number <> 0 Then
        Err.Clear
        For idx = 1 To doomedCount
            hostSheet.Shapes(doomedNames(idx)).Delete
        Next idx
    End If
    Err.Clear
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
    Dim blinkMs As Long

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
            TrackToastParts hostSheet
        Case KIND_INPUT, KIND_COMBO
            If mInputDueMs > 0 Then
                If GetTickCount64() >= mInputDueMs Then FireInput OwnerApp
            End If
            ' Insertion bar blink for the focused float field, at the
            ' Windows rate; a caret set not to blink stays on.
            blinkMs = ReDimUI.CaretBlinkMs
            If blinkMs > 0 Then
                mCaretAccumMs = mCaretAccumMs + elapsedMs
                If mCaretAccumMs >= blinkMs Then
                    mCaretAccumMs = 0
                    mCaretOn = Not mCaretOn
                    Touch
                End If
            ElseIf Not mCaretOn Then
                mCaretOn = True
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

' Amber for warnings: dark on light surfaces and bright on dark ones, so
' it keeps its contrast in every theme without a slot of its own.
Public Property Get WarningColor() As Long
    If RelativeLuminance(mThemeSurface) < 0.2 Then
        WarningColor = RGB(252, 200, 60)
    Else
        WarningColor = RGB(157, 93, 0)
    End If
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
