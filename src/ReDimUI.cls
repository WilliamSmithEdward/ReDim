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
' ReDim 1.0.0 (2026-09-23)
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
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal pauseMs As Long)
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
Private Declare PtrSafe Function RegGetValueW Lib "advapi32" ( _
    ByVal rootKey As LongPtr, ByVal subKeyPath As LongPtr, _
    ByVal entryName As LongPtr, ByVal typeFlags As Long, _
    ByVal typeOut As LongPtr, ByVal dataOut As LongPtr, _
    ByVal sizeOut As LongPtr) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const VK_LEFT_BUTTON As Long = 1
Private Const VK_RIGHT_BUTTON As Long = 2
Private Const SPI_GETCLIENTAREAANIMATION As Long = &H1042
Private Const SPI_GETKEYBOARDDELAY As Long = &H16
Private Const SPI_GETKEYBOARDSPEED As Long = &HA
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE_ZEROED As Long = &H42
Private Const REG_ROOT_USER As Long = &H80000001
Private Const REG_ROOT_MACHINE As Long = &H80000002
Private Const REG_TYPE_DWORD As Long = &H10
Private Const REG_TYPE_ANY As Long = &HFFFF&
' Apps that follow the Windows look compare it at most this often.
Private Const SYSTEM_LOOK_CHECK_MS As Long = 2000
' The command palette's parts; the mdl_ prefix keeps a click on them
' from ending the field's focus.
Private Const PALETTE_CARD_ID As String = "mdl_pal_card"
Private Const PALETTE_FIELD_ID As String = "mdl_pal_field"

Private Const SHAPE_PREFIX As String = "rdm_"
Private Const PART_SEPARATOR As String = "__"
Private Const ERROR_BASE As Long = vbObjectError + 4200
Private Const ERROR_INVALID_ARGUMENT As Long = ERROR_BASE + 1
Private Const ERROR_INVALID_OPERATION As Long = ERROR_BASE + 2
Private Const ERROR_NOT_FOUND As Long = ERROR_BASE + 3
Private Const CLICK_DEBOUNCE_MS As Long = 150
Private Const UNMAPPABLE_POINTS As Double = -1E+18

Private Const ROLE_APP As Long = 1
Private Const ROLE_COMPONENT As Long = 2
Private Const ROLE_THEME As Long = 3
Private Const ROLE_ASYNC_OP As Long = 4
Private Const ROLE_JOB As Long = 5

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
Private Const KIND_SKELETON As Long = 22
Private Const KIND_TABS As Long = 23
Private Const KIND_DATE As Long = 24
Private Const KIND_TABLE As Long = 25
Private Const KIND_BADGE As Long = 26
Private Const KIND_STACK As Long = 27
Private Const KIND_EXPANDER As Long = 28
Private Const KIND_SPARKLINE As Long = 29

' Option rows a SelectBox or ComboBox drop list draws at once unless
' ListRows says otherwise; longer lists window behind pager rows.
Private Const DEFAULT_LIST_ROWS As Long = 8

' List rows open with a gutter for the check that marks the current or
' selected item: the glyph, a tab, then the item, aligned by a tab stop
' this many points in.
Private Const LIST_GUTTER As Double = 14
' Tooltips: the widest a tip grows before it wraps, the tallest control a
' tip sits clear below rather than under the pointer, and how far the
' pointer may drift, in points, and still count as resting.
Private Const TIP_MAX_WIDTH As Double = 240
Private Const TIP_BELOW_MAX_H As Double = 48
Private Const TIP_REST_SLOP As Double = 3

' Transfer paging arrows are square pointer targets of this many points,
' 24 pixels at 96 DPI: the WCAG 2.5.8 minimum target size.
Private Const PAGER_TARGET As Double = 18

' Toast size, in points: the width a toast keeps unless MinWidth and
' MaxWidth let it fit its words, the height of one line, the greatest
' height before the message ends in an ellipsis, and the gap between
' stacked toasts.
Private Const TOAST_WIDTH As Double = 240
Private Const TOAST_MIN_HEIGHT As Double = 40
Private Const TOAST_MAX_HEIGHT As Double = 160
Private Const TOAST_GAP As Double = 6
' The card's corner radius, and the inset it gives the words on each side.
Private Const TOAST_CORNER As Double = 8
Private Const TOAST_TEXT_INSET As Double = 2.35

' Transfer list layout, in points: the header over each panel, the move
' buttons between the panels and the gap around them, and a row.
Private Const TRANSFER_HEADER_H As Double = 18
Private Const TRANSFER_BUTTON_W As Double = 26
Private Const TRANSFER_BUTTON_H As Double = 22
Private Const TRANSFER_GAP As Double = 6
Private Const TRANSFER_ROW_H As Double = 20

' Date picker calendar layout, in points: a day cell, the padding inside
' the card, the month row and the weekday row, and the card they make.
Private Const CAL_CELL_W As Double = 26
Private Const CAL_CELL_H As Double = 22
Private Const CAL_PAD As Double = 6
Private Const CAL_HEAD_H As Double = 24
Private Const CAL_WEEK_H As Double = 16
Private Const CAL_WIDTH As Double = 194
Private Const CAL_DEPTH As Double = 184
' Calendar state slots after the card, header, arrows, and weekdays (1-5):
' the day marks, then one per week.
Private Const CAL_SLOT_HELD As Long = 6
Private Const CAL_SLOT_HOVER As Long = 7
Private Const CAL_SLOT_TODAY As Long = 8
Private Const CAL_SLOT_CURSOR As Long = 9
Private Const CAL_SLOT_WEEK1 As Long = 10
Private Const CAL_SLOTS As Long = 15

' Table cells keep this many points clear at each side of their text.
Private Const TABLE_CELL_PAD As Double = 6

' The dot on a sparkline's last value, in points; the line keeps half
' of it clear at every edge so the dot stays inside the rectangle.
Private Const SPARK_DOT As Double = 5

' The corner radius of a list row's hover and selection fill, a tab's
' hover fill, and an expander header's, in points.
Private Const ROW_CORNER As Double = 4

' How long after the release of a press off a field or open list the
' close waits for Excel to deliver a click to the control itself.
Private Const OFF_PRESS_GRACE_MS As Long = 120

' A gliding knob or tab bar keeps this share of its distance per 50 ms,
' so a glide across a switch settles in about 200 ms at any frame rate.
Private Const GLIDE_DECAY As Double = 0.4

' Sentinel for "the source could not be loaded"; contains characters no
' real path can, so it never collides with one.
Private Const IMAGE_MISSING As String = "?missing?"

' The words ReDim draws or announces on its own, by index. SetUIText
' replaces them by name; UI_TEXT_TABLE holds each name and default.
Private Const UI_NO_ROWS As Long = 1
Private Const UI_NO_ROWS_MATCH As Long = 2
Private Const UI_NO_ITEMS As Long = 3
Private Const UI_NO_MATCHES As Long = 4
Private Const UI_ROW_RANGE As Long = 5
Private Const UI_ROW_RANGE_EMPTY As Long = 6
Private Const UI_FILTER_NOTE As Long = 7
Private Const UI_SELECT_ALL As Long = 8
Private Const UI_SELECT_ALL_FILTERED As Long = 9
Private Const UI_AVAILABLE As Long = 10
Private Const UI_SELECTED As Long = 11
Private Const UI_PANEL_FILTERED As Long = 12
Private Const UI_MORE_ITEMS As Long = 13
Private Const UI_TYPE_COMMAND As Long = 14
Private Const UI_WORKING As Long = 15
Private Const UI_REQUIRED As Long = 16
Private Const UI_OK As Long = 17
Private Const UI_CANCEL As Long = 18
Private Const UI_COLUMN As Long = 19
Private Const UI_IMAGE_NOT_FOUND As Long = 20
Private Const UI_ALT_BUTTON As Long = 21
Private Const UI_ALT_BUSY As Long = 22
Private Const UI_ALT_PROGRESS As Long = 23
Private Const UI_ALT_SPINNER As Long = 24
Private Const UI_ALT_SWITCH_ON As Long = 25
Private Const UI_ALT_SWITCH_OFF As Long = 26
Private Const UI_ALT_EDIT_FIELD As Long = 27
Private Const UI_ALT_MASKED As Long = 28
Private Const UI_ALT_COMBO_BOX As Long = 29
Private Const UI_ALT_MENU_BUTTON As Long = 30
Private Const UI_ALT_DROP_DOWN As Long = 31
Private Const UI_ALT_CHECKED As Long = 32
Private Const UI_ALT_NOT_CHECKED As Long = 33
Private Const UI_ALT_RADIO_GROUP As Long = 34
Private Const UI_ALT_ITEM_SELECTED As Long = 35
Private Const UI_ALT_STEPPER As Long = 36
Private Const UI_ALT_SLIDER As Long = 37
Private Const UI_ALT_TRANSFER_LIST As Long = 38
Private Const UI_ALT_CHECK_LIST As Long = 39
Private Const UI_ALT_IMAGE As Long = 40
Private Const UI_ALT_LOADING As Long = 41
Private Const UI_ALT_EXPANDED As Long = 42
Private Const UI_ALT_COLLAPSED As Long = 43
Private Const UI_ALT_TABLE As Long = 44
Private Const UI_ALT_TABLE_ONE_ROW As Long = 45
Private Const UI_ALT_SORTED_ASCENDING As Long = 46
Private Const UI_ALT_SORTED_DESCENDING As Long = 47
Private Const UI_ALT_SHOWN As Long = 48
Private Const UI_ALT_DATE_PICKER As Long = 49
Private Const UI_ALT_TABS As Long = 50
Private Const UI_ALT_TAB_SELECTED As Long = 51
Private Const UI_ALT_UNAVAILABLE As Long = 52
Private Const UI_ALT_SPARKLINE As Long = 53
Private Const UI_ALT_SPARKLINE_EMPTY As Long = 54
Private Const UI_TEXT_COUNT As Long = 54

' Each UI text's name and English default, "name=words", in index
' order and joined by bars; EnsureUIWords splits it once. One constant
' rather than a Select Case: VBA pays for a procedure's first call by
' its size, and this table sits on every first draw.
Private Const UI_TEXT_TABLE As String = _
    "NoRows=No rows|NoRowsMatch=No rows match|NoItems=No items|NoMatches=No matches|" & _
    "RowRange={0}-{1} of {2}|RowRangeEmpty=0 of 0|FilterNote=Filter ""{0}"": {1}|" & _
    "SelectAll=Select all ({0}/{1})|SelectAllFiltered=Select all ""{0}"" ({1}/{2})|" & _
    "Available=Available|Selected=Selected|PanelFiltered={0} ""{1}"" ({2} of {3})|" & _
    "MoreItems={0} more|TypeCommand=Type a command|Working=Working...|" & _
    "Required=Required|OK=OK|Cancel=Cancel|Column=Column {0}|" & _
    "ImageNotFound=(image not found)|AltButton={0}, button|AltBusy=busy|" & _
    "AltProgress=Progress, {0} percent|AltSpinner=Busy|AltSwitchOn=Switch, on|" & _
    "AltSwitchOff=Switch, off|AltEditField=Edit field|AltMasked=masked|" & _
    "AltComboBox=Combo box|AltMenuButton={0}, menu button|AltDropDown=Drop-down, {0}|" & _
    "AltChecked={0}, checkbox, checked|AltNotChecked={0}, checkbox, not checked|" & _
    "AltRadioGroup=Radio group|AltItemSelected={0} selected|AltStepper=Stepper, {0}|" & _
    "AltSlider=Slider, {0}|AltTransferList=Transfer list|" & _
    "AltCheckList=Checklist, {0} of {1} checked|AltImage=Image|AltLoading=Loading|" & _
    "AltExpanded={0}, expander, expanded|AltCollapsed={0}, expander, collapsed|" & _
    "AltTable=Table, {0} rows|AltTableOneRow=Table, 1 row|" & _
    "AltSortedAscending=sorted by {0}, ascending|" & _
    "AltSortedDescending=sorted by {0}, descending|AltShown={0} shown|" & _
    "AltDatePicker=Date picker, {0}|AltTabs=Tabs|" & _
    "AltTabSelected={0} selected, tab {1} of {2}|AltUnavailable=unavailable|" & _
    "AltSparkline=Trend, {0} values, low {1}, high {2}, last {3}|" & _
    "AltSparklineEmpty=Trend, no values"

Private Const VARIANT_PRIMARY As Long = 0
Private Const VARIANT_SECONDARY As Long = 1
Private Const VARIANT_SUCCESS As Long = 2
Private Const VARIANT_DANGER As Long = 3
Private Const VARIANT_NEUTRAL As Long = 4
Private Const VARIANT_WARNING As Long = 5

Private Const REL_NONE As Long = 0
Private Const REL_BELOW As Long = 1
Private Const REL_RIGHT_OF As Long = 2
Private Const REDIM_VERSION As String = "1.0.0"

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
' The Alt+letter chords bound for the active sheet's access keys.
Private mBoundAccessKeys As String
' The shortcut codes bound now, each followed by Chr$(1).
Private mBoundShortcuts As String
' Test seam: a fixed pointer position, in sheet points, that stands in
' for the cursor until cleared.
Private mPointerOverridden As Boolean
Private mPointerX As Double
Private mPointerY As Double
' The override's left button, and the pointer as read once per pump
' frame: 0 not read yet this frame, 1 read, 2 unmappable.
Private mPointerButtonHeld As Boolean
Private mFramePointerState As Long
Private mFramePointerX As Double
Private mFramePointerY As Double
' The caret blink interval Windows reported when a field last took
' focus; 0 when the user turned blinking off.
Private mCaretBlinkMs As Long
' Motion preference: follow Windows (0), reduced (1), or full (2); and
' the Windows setting as last read, with when.
Private mMotionChoice As Long
Private mMotionOffWindows As Boolean
Private mMotionReadMs As LongLong
' The UI text table: each entry's name and the words it holds now.
Private mUITextNames() As String
Private mUIWords() As String
Private mUIWordsReady As Boolean
' Guarded-swallow observability: every silently trapped tick or key
' dispatch fault counts here, so tests can prove a run was clean.
Private mTickFaults As Long
' The Windows icon font, found once. The Windows look ThemeSystem reads,
' overridden as a test seam (0 reads Windows, 1 light, 2 dark) with its
' accent (-1 for none), and when the apps that follow it last compared.
Private mIconFontName As String
Private mSystemLookOverride As Long
Private mSystemAccentOverride As Long
Private mSystemLookCheckedMs As LongLong
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
' The state keys in the order they were first set; a Collection keeps
' no key list of its own.
Private mStateKeys As Collection
' Whether the app's sheet held none of its shapes when it mounted, and
' the shape names it has drawn since. On such a sheet a name it never
' drew cannot exist, so the lookup that would miss is skipped: a missed
' Shapes(name) scans the whole sheet and raises, 0.1 ms at 100 shapes
' and 0.8 ms at 1,000.
Private mSurfaceFresh As Boolean
Private mDrawnShapes As Collection
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
' Set when a toast's height changed during a flush; the tray restacks
' once the flush ends.
Private mToastRestackPending As Boolean
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
' Whether the app follows the Windows look, and the look it last took.
Private mFollowSystem As Boolean
Private mAppliedSystemLook As String
' Stacks whose members' flow changed during a flush, laid out again when
' the flush ends.
Private mPendingReflows As Collection
' The command palette: whether it is on, what each entry it lists runs,
' and the control that had the keys when it opened.
Private mPaletteOn As Boolean
Private mPaletteActions As Collection
Private mFocusBeforePalette As String

' Component fields.
Private mComponentId As String
Private mOwnerAppId As String
Private mKind As Long
Private mAnchor As String
Private mLeft As Double
Private mTop As Double
Private mWidth As Double
Private mHeight As Double
Private mText As String
Private mBusyText As String
Private mBusyTextGiven As Boolean
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
' The action button's border, -1 for the theme's accent.
Private mToastActionBorder As Long
' A toast's size limits (0 for the defaults) and the tray top its stack
' measures from.
Private mToastMinWidth As Double
Private mToastMaxWidth As Double
Private mToastMaxHeight As Double
Private mToastOriginTop As Double
Private mSlideRemaining As Double
Private mToastExiting As Boolean
Private mExitRemaining As Double
' The countdown a toast keeps while the pointer rests on it; 0 when free.
Private mToastHeldMs As LongLong
' Pointer effects, per app: whether hover and press looks draw, the
' control and part under the pointer and when the pointer got there,
' the press look shown, where the left button went down, and the last
' pointer position, which tells a moving pointer from a resting one.
Private mPointerEffects As Boolean
Private mHoverTargetId As String
Private mHoverTargetPart As String
Private mHoverPressedShown As Boolean
Private mPressTargetId As String
Private mPressTargetPart As String
Private mHoverPrevDown As Boolean
Private mHoverPointX As Double
Private mHoverPointY As Double
' A control's own pointer look: hovered, the part hovered, and pressed;
' and the transfer button or arrow tinted for it.
Private mHovered As Boolean
Private mHoveredPart As String
Private mHoverPressed As Boolean
Private mAppliedPartTint As String
' Tooltips: a control's note for a resting pointer and its note while
' disabled; per app, whether any control has one, the control the
' tooltip delay runs for and since when, the control showing one, and
' whether a press put tooltips away until the pointer leaves.
Private mTooltip As String
Private mDisabledReason As String
Private mHasTooltips As Boolean
Private mTipTargetId As String
Private mTipSinceMs As LongLong
Private mTipShownId As String
Private mTipSuppressed As Boolean
' Where the pointer came to rest, and the shown tip's edges, in points.
Private mTipRestX As Double
Private mTipRestY As Double
Private mTipLeft As Double
Private mTipTop As Double
Private mTipRight As Double
Private mTipBottom As Double
' Hold-to-repeat, per app: the control and part a held press repeats,
' when the next repeat is due, how many ran, and the button last frame;
' per control, the value a hold started from and the release click it
' swallows.
Private mHoldId As String
Private mHoldPart As String
Private mHoldNextMs As LongLong
Private mHoldRepeats As Long
Private mHoldPrevDown As Boolean
Private mHoldValueAtStart As Double
Private mHoldSuppressUntilMs As LongLong
' Transfer gestures: the last row click, which tells a double click; a
' press-drag from a row, its panel (0 left, 1 right) and item, what it
' became (1 a range selection, 2 a drag onto the other panel), and the
' button last frame; and the panel drawn as the drop target (1 left,
' 2 right, 0 none).
Private mLastRowClickMs As LongLong
Private mLastRowClickPart As String
Private mRowDragOn As Boolean
Private mRowDragPanel As Long
Private mRowDragFrom As Long
Private mRowDragMode As Long
Private mRowDragPrevDown As Boolean
Private mDropPanel As Long
Private mAppliedDropPanel As Long
' Transfer reorder: whether the chosen panel shows its up and down
' arrows, and the signature they were drawn with.
Private mReorderable As Boolean
Private mAppliedReorderSig As String
' Type-to-filter while a transfer or check list has the keys: each
' transfer panel's filter and the check list's; a counter every item
' change bumps; and the item numbers each filtered transfer panel
' shows, with the key they were read under.
Private mFilterLeft As String
Private mFilterRight As String
Private mFilterCheck As String
Private mItemsVersion As Long
Private mShownLeft() As Long
Private mShownRight() As Long
Private mShownLeftCount As Long
Private mShownRightCount As Long
Private mShownLeftKey As String
Private mShownRightKey As String
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
' Item values by item text, for items that carry one: WritesTo writes
' the value and ItemValueAt reads it. Keyed by text, without case, so a
' value follows its item between transfer panels.
Private mItemValueMap As Collection
' SelectBox items that cannot be picked: disabled ones and group
' headers, as sets of item numbers that inserts and removals keep in
' step.
Private mDisabledItems As Collection
Private mGroupHeaders As Collection
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
' Field adornments: a caption above the control, a hint below it (a
' message takes its place), a required mark with the message an empty
' commit shows, and an error set from outside; with the signatures the
' parts were drawn with.
Private mCaptionText As String
Private mHintText As String
Private mRequired As Boolean
Private mRequiredMessage As String
Private mErrorText As String
Private mAppliedCaptionSig As String
Private mAppliedHintSig As String
' Skeleton: the text lines it stands in for (0 draws one block), the
' time into its pulse, and the pulse step drawn.
Private mSkeletonLines As Long
Private mPulseMs As Double
Private mAppliedPulseStep As Long
' Tabs: each tab's right edge as drawn, measured from the control's
' left, for the pointer; and the tab its panels were last shown for
' (-2 before the first sync).
Private mTabRights() As Double
Private mAppliedPanelTab As Long
' A component on a Tabs panel: the Tabs control and the tab it belongs
' to, whether that tab hides it now, and the visibility it has apart
' from the tab, which Visible and BindVisible set while it is hidden.
Private mTabOwnerId As String
Private mTabNumber As Long
Private mTabShut As Boolean
Private mVisibleBeforeTab As Boolean
' Date picker: the month its calendar shows (the first of it), the day
' the keys have reached, the earliest and latest days it picks (0 for
' an open side), and the face's Format$ pattern (empty for the system's
' short date). The date itself is mValue, 0 while none is picked.
Private mCalendarMonth As Date
Private mCalendarCursor As Date
Private mDateEarliest As Date
Private mDateLatest As Date
Private mDateFormat As String
' Table: cells by column, then row, so rows can grow; each column's
' header, Format$ pattern, asked width (0 fits the content), and whether
' it holds numbers; the columns and rows held and a count of data
' changes; the order rows show in, as source row numbers, and whether it
' is current; the column sorted by (0 for the source order) and its
' direction; the rows scrolled past the top; and what the last draw
' laid out: each column's right edge from the table's left, with the
' key it was measured under, the row pitch, the row slots, whether the
' footer shows, and each part's look.
Private mTableCells() As Variant
Private mTableHeads() As String
Private mColumnFormats() As String
Private mColumnPoints() As Double
Private mColumnNumeric() As Boolean
Private mTableCols As Long
Private mTableRows As Long
Private mTableVersion As Long
Private mTableOrder() As Long
Private mTableOrderCurrent As Boolean
Private mSortColumn As Long
Private mSortDescending As Boolean
Private mTableOffset As Long
Private mColumnEdges() As Double
Private mColumnLayoutKey As String
Private mTableRowPitch As Double
Private mTableSlots As Long
Private mTableFooterOn As Boolean
Private mAppliedTableRowStates() As String
Private mAppliedTableHeadStates() As String
Private mAppliedTableFootSig As String
Private mAppliedTableColsDrawn As Long
Private mAppliedTableSlotsDrawn As Long
' A table's typed filter, the rows it shows under it in order (source
' row numbers) with the key they were found under, a count of order
' rebuilds, the empty row's words, and the handler a row's double
' click or Enter runs.
Private mTableFilter As String
Private mTableShown() As Long
Private mTableShownCount As Long
Private mTableShownKey As String
Private mTableOrderGen As Long
Private mTableEmptyText As String
Private mRowOpenProc As String
' A float TextInput that shows dots for its characters.
Private mMasked As Boolean
' A press judged off the field or open list, waiting for its release,
' and when the release was seen.
Private mOffPressPending As Boolean
Private mOffPressUpMs As LongLong
' A part gliding to its place, a switch's knob or a tab strip's bar: the
' part, where it is headed, and how far it still is from there, which
' the frames close. The state and tab it last glided from.
Private mGlidePart As String
Private mGlideTargetLeft As Double
Private mGlideTargetWidth As Double
Private mGlideGapLeft As Double
Private mGlideGapWidth As Double
Private mAppliedGlideBase As String
Private mAppliedIndicatorTab As Long
' A sparkline's numbers, a count of their changes, and the state its
' line and dot were drawn for.
Private mSparkValues() As Double
Private mSparkCount As Long
Private mSparkVersion As Long
Private mAppliedSparkSig As String
' An icon before the control's text: its name, its character, and the
' character the text was last styled with.
Private mIconName As String
Private mIconGlyph As String
Private mAppliedIconGlyph As String
' A count or note in a pill on the control's top-right corner, and the
' signature it was drawn with.
Private mBadgeText As String
Private mAppliedBadgeSig As String
' Stack: its members in the order they joined, the gap between them, the
' padding inside it, whether they run across instead of down, and whether
' they take its width; a count of flow changes, and the layout last
' made from them with the count, place, and width it was made under;
' and whether the stack was shown when its members last applied.
Private mStackMembers As Collection
Private mStackGap As Double
Private mStackPadding As Double
Private mStackAcross As Boolean
Private mStackStretch As Boolean
Private mStackGen As Long
Private mStackLaidGen As Long
Private mStackLaidLeft As Double
Private mStackLaidTop As Double
Private mStackLaidWidth As Double
Private mStackSlotIds() As String
Private mStackSlotLefts() As Double
Private mStackSlotTops() As Double
Private mStackSlotSpans() As Double
Private mStackSlotCount As Long
Private mAppliedStackShown As Boolean
Private mStackAutoWidth As Boolean
' A control in a stack: the stack, and the flow it last took there.
Private mStackId As String
Private mAppliedFlowSig As String
' Commands: a MenuButton's rows, or an app's palette entries. Whether a
' select is a menu, each command's handler and icon by its text, the
' texts in the order they came (apps), and the command run last.
Private mIsMenu As Boolean
Private mCommandProcs As Collection
Private mCommandIcons As Collection
Private mCommandOrder As Collection
Private mLastCommand As String
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
' A menu's rows run as wide as its widest command needs; 0 while the rows
' take the face's span.
Private mListLeft As Double
Private mListWidth As Double
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
' Whether the slider's value bubble is drawn.
Private mAppliedBubble As Boolean
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
' The keyboard shortcut that clicks the control, as an OnKey code.
Private mShortcut As String
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
' A sheet coming back to the front re-arms the pump for the watches that
' need it there: a slider's press watch, pointer effects. The pump stops
' itself while no app has work, and leaving the sheet takes theirs.
Private Sub mExcelApp_SheetActivate(ByVal Sh As Object)
    RefreshAccessKeys
    If HasPendingWork() Then RequestPump
    CheckSystemLook
End Sub

Private Sub mExcelApp_WorkbookActivate(ByVal Wb As Workbook)
    RefreshAccessKeys
    If HasPendingWork() Then RequestPump
    CheckSystemLook
End Sub

Private Sub mExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    CheckSystemLook
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
    ' A click on the control: the press that led to it was on it.
    componentValue.CancelOffPress
    If targetPart = "tt" Then
        appValue.DismissTooltip
        Exit Sub
    End If
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
    ' The keypad's decimal key types the separator Excel would.
    If keyText = "{DECIMAL}" Then keyText = DecimalMark()
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
    RefreshShortcuts
End Sub

' Binds each shortcut the active sheet's apps declare and releases the
' rest, as access keys are; a released key goes back to Excel.
Private Sub RefreshShortcuts()
    Dim appValue As ReDimUI
    Dim wantedCodes As String
    Dim listedCodes() As String
    Dim idx As Long

    On Error Resume Next
    For Each appValue In mApps
        If appValue.Sheet Is ActiveSheet Then
            listedCodes = Split(appValue.ShortcutCodes, Chr$(1))
            For idx = 0 To UBound(listedCodes)
                If LenB(listedCodes(idx)) > 0 Then
                    If Not CodeListHas(wantedCodes, listedCodes(idx)) Then
                        wantedCodes = CodeListAdd(wantedCodes, listedCodes(idx))
                    End If
                End If
            Next idx
        End If
    Next appValue
    listedCodes = Split(mBoundShortcuts, Chr$(1))
    For idx = 0 To UBound(listedCodes)
        If LenB(listedCodes(idx)) > 0 Then
            If Not CodeListHas(wantedCodes, listedCodes(idx)) Then
                Application.OnKey listedCodes(idx)
            End If
        End If
    Next idx
    listedCodes = Split(wantedCodes, Chr$(1))
    For idx = 0 To UBound(listedCodes)
        If LenB(listedCodes(idx)) > 0 Then
            If Not CodeListHas(mBoundShortcuts, listedCodes(idx)) Then
                BindShortcut listedCodes(idx)
            End If
        End If
    Next idx
    On Error GoTo 0
    mBoundShortcuts = wantedCodes
End Sub

' A focused field's release gives its keys back to Excel, and a shortcut
' can share one (Ctrl+A, say, which the field took for select-all); the
' bound shortcuts take their keys again.
Private Sub RebindShortcuts()
    Dim listedCodes() As String
    Dim idx As Long

    If LenB(mBoundShortcuts) = 0 Then Exit Sub
    listedCodes = Split(mBoundShortcuts, Chr$(1))
    On Error Resume Next
    For idx = 0 To UBound(listedCodes)
        If LenB(listedCodes(idx)) > 0 Then BindShortcut listedCodes(idx)
    Next idx
    On Error GoTo 0
End Sub

Private Sub BindShortcut(ByVal keyCode As String)
    Application.OnKey keyCode, "'RdxShortcut """ & keyCode & """'"
End Sub

' Lists of shortcut codes: each code closed in Chr$(1) marks, so "{F5}"
' is not found inside "+{F5}", compared without regard to case.
Private Function CodeListHas(ByVal codeList As String, ByVal keyCode As String) As Boolean
    CodeListHas = (InStr(1, codeList, Chr$(1) & keyCode & Chr$(1), vbTextCompare) > 0)
End Function

Private Function CodeListAdd(ByVal codeList As String, ByVal keyCode As String) As String
    If LenB(codeList) = 0 Then codeList = Chr$(1)
    CodeListAdd = codeList & keyCode & Chr$(1)
End Function

' A shortcut: the first app on the active sheet with a control declaring
' it clicks that control.
Friend Sub DispatchShortcut(ByVal keyCode As String)
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.DispatchShortcut keyCode
        Exit Sub
    End If
    EnsureFactory
    For Each appValue In mApps
        If appValue.Sheet Is ActiveSheet Then
            If appValue.ActivateShortcut(keyCode) Then Exit Sub
        End If
    Next appValue
End Sub

' Test seam: the shortcut codes bound now, in their Chr$(1)-marked list.
Friend Property Get BoundShortcuts() As String
    If Not IsFactory(Me) Then
        BoundShortcuts = ReDimUI.BoundShortcuts
        Exit Property
    End If
    BoundShortcuts = mBoundShortcuts
End Property

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
    If Not mFocusHandoff Then
        RdxReleaseKeys
        RebindShortcuts
    End If
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

' Test seam: pointer-driven behavior (outside presses, drags, holds, and
' hover) reads this sheet position, with buttonDown as the left button,
' instead of the cursor and the mouse until ClearPointerOverride.
Friend Sub OverridePointer( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double, _
    Optional ByVal buttonDown As Boolean = False _
)
    If Not IsFactory(Me) Then
        ReDimUI.OverridePointer pointsX, pointsY, buttonDown
        Exit Sub
    End If
    mPointerOverridden = True
    mPointerX = pointsX
    mPointerY = pointsY
    mPointerButtonHeld = buttonDown
    mFramePointerState = 0
End Sub

Friend Sub ClearPointerOverride()
    If Not IsFactory(Me) Then
        ReDimUI.ClearPointerOverride
        Exit Sub
    End If
    mPointerOverridden = False
    mPointerButtonHeld = False
    mFramePointerState = 0
End Sub

' The left button, and whether either button is down, from the override
' while one is set.
Friend Sub ReadButtons(ByRef leftDown As Boolean, ByRef eitherDown As Boolean)
    If Not IsFactory(Me) Then
        ReDimUI.ReadButtons leftDown, eitherDown
        Exit Sub
    End If
    If mPointerOverridden Then
        leftDown = mPointerButtonHeld
        eitherDown = leftDown
    Else
        leftDown = (GetAsyncKeyState(VK_LEFT_BUTTON) < 0)
        eitherDown = leftDown Or (GetAsyncKeyState(VK_RIGHT_BUTTON) < 0)
    End If
End Sub

' The pointer in sheet points, read at most once per pump frame and
' shared by every watch that asks for it. False when it cannot be mapped.
Friend Function FramePointer(ByRef pointsX As Double, ByRef pointsY As Double) As Boolean
    Dim screenPoint As RdxScreenPoint

    If Not IsFactory(Me) Then
        FramePointer = ReDimUI.FramePointer(pointsX, pointsY)
        Exit Function
    End If
    If mFramePointerState = 0 Then
        mFramePointerState = 2
        If mPointerOverridden Then
            mFramePointerX = mPointerX
            mFramePointerY = mPointerY
            mFramePointerState = 1
        ElseIf GetCursorPos(screenPoint) <> 0 Then
            mFramePointerX = ScreenXToSheetPoints(screenPoint.pixelX)
            mFramePointerY = ScreenYToSheetPoints(screenPoint.pixelY)
            If mFramePointerX <> UNMAPPABLE_POINTS _
                And mFramePointerY <> UNMAPPABLE_POINTS Then
                mFramePointerState = 1
            End If
        End If
    End If
    If mFramePointerState = 1 Then
        pointsX = mFramePointerX
        pointsY = mFramePointerY
        FramePointer = True
    End If
End Function

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
    mFramePointerState = 0
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
    If runWork Then
        On Error Resume Next
        CheckSystemLook
        If Err.Number <> 0 Then
            Err.Clear
            NoteTickFault
        End If
        On Error GoTo 0
    End If
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
' set, else the Windows setting, read again at most once a second since
' users change it while frames ask many times a second.
Public Property Get MotionReduced() As Boolean
    Dim animationsOn As Long
    Dim nowMs As LongLong

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
            nowMs = GetTickCount64()
            If mMotionReadMs = 0 Or nowMs - mMotionReadMs >= 1000 Then
                animationsOn = 1
                mMotionOffWindows = False
                If SystemParametersInfo(SPI_GETCLIENTAREAANIMATION, 0, _
                    animationsOn, 0) <> 0 Then
                    mMotionOffWindows = (animationsOn = 0)
                End If
                mMotionReadMs = nowMs
            End If
            MotionReduced = mMotionOffWindows
    End Select
End Property

' Replaces words ReDim draws or announces on its own, such as "No rows"
' or ", button" for screen readers, for an app in another language;
' UITextKeys names them. "{0}", "{1}", and on in the words take the
' numbers and names of the moment. Controls take new words as they
' repaint, and Render repaints an app at once.
Public Sub SetUIText(ByVal textKey As String, ByVal wordsText As String)
    Dim textIndex As Long

    If Not IsFactory(Me) Then
        ReDimUI.SetUIText textKey, wordsText
        Exit Sub
    End If
    textIndex = UITextIndex(textKey, "SetUIText")
    mUIWords(textIndex) = wordsText
End Sub

' The words a UI text key holds now.
Public Property Get UIText(ByVal textKey As String) As String
    If Not IsFactory(Me) Then
        UIText = ReDimUI.UIText(textKey)
        Exit Property
    End If
    UIText = mUIWords(UITextIndex(textKey, "UIText"))
End Property

' Every UI text key, in order, for a translation table.
Public Property Get UITextKeys() As Variant
    If Not IsFactory(Me) Then
        UITextKeys = ReDimUI.UITextKeys
        Exit Property
    End If
    EnsureUIWords
    UITextKeys = mUITextNames
End Property

' Puts every UI text back to its English default.
Public Sub ResetUIText()
    If Not IsFactory(Me) Then
        ReDimUI.ResetUIText
        Exit Sub
    End If
    mUIWordsReady = False
    EnsureUIWords
End Sub

Friend Function UIWords(ByVal textIndex As Long) As String
    If Not mUIWordsReady Then EnsureUIWords
    UIWords = mUIWords(textIndex)
End Function

Private Sub EnsureUIWords()
    Dim textIndex As Long
    Dim tableEntries() As String
    Dim definedAs As String
    Dim splitAt As Long

    If mUIWordsReady Then Exit Sub
    tableEntries = Split(UI_TEXT_TABLE, "|")
    ReDim mUITextNames(1 To UI_TEXT_COUNT)
    ReDim mUIWords(1 To UI_TEXT_COUNT)
    For textIndex = 1 To UI_TEXT_COUNT
        definedAs = tableEntries(textIndex - 1)
        splitAt = InStr(definedAs, "=")
        mUITextNames(textIndex) = Left$(definedAs, splitAt - 1)
        mUIWords(textIndex) = Mid$(definedAs, splitAt + 1)
    Next textIndex
    mUIWordsReady = True
End Sub

' A UI text key's index, in any case; raises for a name that is none.
Private Function UITextIndex(ByVal textKey As String, ByVal memberName As String) As Long
    Dim textIndex As Long

    EnsureUIWords
    For textIndex = 1 To UI_TEXT_COUNT
        If StrComp(mUITextNames(textIndex), textKey, vbTextCompare) = 0 Then
            UITextIndex = textIndex
            Exit Function
        End If
    Next textIndex
    RaiseError ERROR_INVALID_ARGUMENT, memberName, _
        "No UI text is named '" & textKey & "'; UITextKeys lists the names."
End Function

' The UI text at textIndex with each "{n}" in it replaced by the nth
' fill, counting from 0. One pass, so braces a fill brings stay as they
' are.
Private Function Phrase(ByVal textIndex As Long, ParamArray fillIns() As Variant) As String
    Dim phraseText As String
    Dim scanFrom As Long
    Dim openAt As Long
    Dim closeAt As Long
    Dim fillToken As String

    phraseText = ReDimUI.UIWords(textIndex)
    scanFrom = 1
    Do
        openAt = InStr(scanFrom, phraseText, "{")
        If openAt = 0 Then Exit Do
        closeAt = InStr(openAt, phraseText, "}")
        If closeAt = 0 Then Exit Do
        fillToken = Mid$(phraseText, openAt + 1, closeAt - openAt - 1)
        Phrase = Phrase & Mid$(phraseText, scanFrom, openAt - scanFrom)
        If fillToken Like "#" And CLng(Val(fillToken)) <= UBound(fillIns) Then
            Phrase = Phrase & CStr(fillIns(CLng(Val(fillToken))))
        Else
            Phrase = Phrase & Mid$(phraseText, openAt, closeAt - openAt + 1)
        End If
        scanFrom = closeAt + 1
    Loop
    Phrase = Phrase & Mid$(phraseText, scanFrom)
End Function

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

' The Windows look: the dark theme while Windows apps use dark mode and
' the light one otherwise, with the Windows accent color as the primary,
' moved until it reads, 3:1 against the surface and 4.5:1 under its ink.
Public Function ThemeSystem() As ReDimUI
    Dim darkOn As Boolean
    Dim accentRgb As Long
    Dim themeValue As ReDimUI

    If Not IsFactory(Me) Then
        Set ThemeSystem = ReDimUI.ThemeSystem
        Exit Function
    End If
    ReadSystemLook darkOn, accentRgb
    If darkOn Then
        Set themeValue = ThemeDark
    Else
        Set themeValue = ThemeLight
    End If
    If accentRgb >= 0 Then
        themeValue.WithPrimary ReadableAccent(accentRgb, themeValue.SurfaceColor, _
            themeValue.OnPrimaryColor, darkOn), themeValue.OnPrimaryColor
    End If
    Set ThemeSystem = themeValue
End Function

' The accent moved toward black on a light surface, or white on a dark
' one, a tenth at a time, until it clears 3:1 against the surface and
' 4.5:1 under its ink.
Private Function ReadableAccent( _
    ByVal accentRgb As Long, _
    ByVal surfaceRgb As Long, _
    ByVal inkRgb As Long, _
    ByVal darkOn As Boolean _
) As Long
    Dim stepNo As Long
    Dim towardRgb As Long

    If darkOn Then
        towardRgb = RGB(255, 255, 255)
    Else
        towardRgb = RGB(0, 0, 0)
    End If
    ReadableAccent = accentRgb
    For stepNo = 1 To 20
        If ContrastRatio(ReadableAccent, surfaceRgb) >= 3 _
            And ContrastRatio(inkRgb, ReadableAccent) >= 4.5 Then Exit Function
        ReadableAccent = BlendRgb(ReadableAccent, towardRgb, 0.1)
    Next stepNo
End Function

' Dark mode and the accent from the user's Windows settings: light with no
' accent where the registry holds neither.
Private Sub ReadSystemLook(ByRef darkOn As Boolean, ByRef accentRgb As Long)
    Dim lightFlag As Long
    Dim accentValue As Long

    If mSystemLookOverride <> 0 Then
        darkOn = (mSystemLookOverride = 2)
        accentRgb = mSystemAccentOverride
        Exit Sub
    End If
    darkOn = False
    If RegistryDword(REG_ROOT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", _
        "AppsUseLightTheme", lightFlag) Then darkOn = (lightFlag = 0)
    accentRgb = -1
    If RegistryDword(REG_ROOT_USER, "Software\Microsoft\Windows\DWM", _
        "AccentColor", accentValue) Then accentRgb = accentValue And &HFFFFFF
End Sub

Private Function RegistryDword( _
    ByVal rootKey As Long, _
    ByVal subKeyPath As String, _
    ByVal entryName As String, _
    ByRef dwordValue As Long _
) As Boolean
    Dim readValue As Long
    Dim dataSize As Long

    dataSize = 4
    RegistryDword = (RegGetValueW(CLngPtr(rootKey), StrPtr(subKeyPath), _
        StrPtr(entryName), REG_TYPE_DWORD, 0, VarPtr(readValue), VarPtr(dataSize)) = 0)
    If RegistryDword Then dwordValue = readValue
End Function

Private Function RegistryValueExists( _
    ByVal rootKey As Long, _
    ByVal subKeyPath As String, _
    ByVal entryName As String _
) As Boolean
    RegistryValueExists = (RegGetValueW(CLngPtr(rootKey), StrPtr(subKeyPath), _
        StrPtr(entryName), REG_TYPE_ANY, 0, 0, 0) = 0)
End Function

' Test seam: the look ThemeSystem reads in place of Windows until
' cleared, with accentRgb -1 for no accent.
Friend Sub OverrideSystemLook(ByVal darkOn As Boolean, ByVal accentRgb As Long)
    If Not IsFactory(Me) Then
        ReDimUI.OverrideSystemLook darkOn, accentRgb
        Exit Sub
    End If
    If darkOn Then
        mSystemLookOverride = 2
    Else
        mSystemLookOverride = 1
    End If
    mSystemAccentOverride = accentRgb
End Sub

Friend Sub ClearSystemLookOverride()
    If Not IsFactory(Me) Then
        ReDimUI.ClearSystemLookOverride
        Exit Sub
    End If
    mSystemLookOverride = 0
End Sub

' Re-themes the apps that follow the Windows look when it changed, dark
' mode or the accent. The registry is read at most every two seconds
' unless forced, and only while some app follows.
' The palette key: opens the palette of the app on the active sheet.
Friend Sub OpenPaletteOnActiveSheet()
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.OpenPaletteOnActiveSheet
        Exit Sub
    End If
    If mApps Is Nothing Then Exit Sub
    For Each appValue In mApps
        If appValue.PaletteOn Then
            If appValue.Sheet Is ActiveSheet Then
                appValue.OpenCommandPalette
                Exit Sub
            End If
        End If
    Next appValue
End Sub

' The palette field's pick: the app whose palette is open runs it.
Friend Sub RunPalettePick()
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.RunPalettePick
        Exit Sub
    End If
    If mApps Is Nothing Then Exit Sub
    For Each appValue In mApps
        If appValue.PaletteOpen Then
            appValue.RunPaletteEntry
            Exit Sub
        End If
    Next appValue
End Sub

Friend Sub CheckSystemLook(Optional ByVal forceRead As Boolean = False)
    Dim appValue As ReDimUI
    Dim nowMs As LongLong
    Dim lookKey As String
    Dim darkOn As Boolean
    Dim accentRgb As Long

    If Not IsFactory(Me) Then
        ReDimUI.CheckSystemLook forceRead
        Exit Sub
    End If
    If mApps Is Nothing Then Exit Sub
    nowMs = GetTickCount64()
    If Not forceRead Then
        If nowMs - mSystemLookCheckedMs < SYSTEM_LOOK_CHECK_MS Then Exit Sub
    End If
    mSystemLookCheckedMs = nowMs
    For Each appValue In mApps
        If appValue.FollowsSystemTheme Then
            If LenB(lookKey) = 0 Then
                ReadSystemLook darkOn, accentRgb
                lookKey = CStr(darkOn) & ":" & CStr(accentRgb)
            End If
            appValue.MatchSystemLook lookKey
        End If
    Next appValue
End Sub

' The Windows icon font: Segoe Fluent Icons where Windows 11 installed
' it, and Segoe MDL2 Assets, which Windows 10 and 11 both carry,
' otherwise.
Public Property Get IconFont() As String
    If Not IsFactory(Me) Then
        IconFont = ReDimUI.IconFont
        Exit Property
    End If
    If LenB(mIconFontName) = 0 Then
        If RegistryValueExists(REG_ROOT_MACHINE, _
            "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", _
            "Segoe Fluent Icons (TrueType)") Then
            mIconFontName = "Segoe Fluent Icons"
        Else
            mIconFontName = "Segoe MDL2 Assets"
        End If
    End If
    IconFont = mIconFontName
End Property

' The character that draws a named icon in IconFont, for text of your
' own; one character passes through as itself.
Public Function IconGlyph(ByVal iconName As String) As String
    Dim codePoint As Long

    If Len(iconName) = 1 Then
        IconGlyph = iconName
        Exit Function
    End If
    codePoint = IconCodePoint(iconName)
    If codePoint = 0 Then
        RaiseError ERROR_INVALID_ARGUMENT, "IconGlyph", _
            "No icon is named '" & iconName & "'; IconNames lists them."
    End If
    IconGlyph = ChrW(codePoint)
End Function

' Every icon name, in the order the API guide lists them.
Public Function IconNames() As String
    IconNames = "Add, Remove, Delete, Edit, Save, Search, Filter, Sort, Settings, " & _
        "Refresh, Sync, Close, Check, More, Menu, ChevronDown, ChevronUp, " & _
        "ChevronLeft, ChevronRight, ArrowUp, ArrowDown, Back, Forward, Home, " & _
        "Calendar, Clock, Info, Warning, Error, Help, Download, Upload, Share, " & _
        "Send, Mail, Phone, Person, People, Lock, Unlock, Star, StarFill, " & _
        "Heart, HeartFill, Flag, Tag, Pin, View, Copy, Paste, Cut, Undo, Redo, " & _
        "Play, Pause, Stop, Folder, Document, Print, Attach, Link, Globe, " & _
        "Cloud, Camera, Library, Shop, ZoomIn, ZoomOut, List, Keyboard, Power"
End Function

' Code points in Segoe MDL2 Assets, which Segoe Fluent Icons keeps.
Private Function IconCodePoint(ByVal iconName As String) As Long
    Select Case LCase$(iconName)
        Case "add": IconCodePoint = &HE710&
        Case "remove": IconCodePoint = &HE738&
        Case "delete": IconCodePoint = &HE74D&
        Case "edit": IconCodePoint = &HE70F&
        Case "save": IconCodePoint = &HE74E&
        Case "search": IconCodePoint = &HE721&
        Case "filter": IconCodePoint = &HE71C&
        Case "sort": IconCodePoint = &HE8CB&
        Case "settings": IconCodePoint = &HE713&
        Case "refresh": IconCodePoint = &HE72C&
        Case "sync": IconCodePoint = &HE895&
        Case "close": IconCodePoint = &HE711&
        Case "check": IconCodePoint = &HE73E&
        Case "more": IconCodePoint = &HE712&
        Case "menu": IconCodePoint = &HE700&
        Case "chevrondown": IconCodePoint = &HE70D&
        Case "chevronup": IconCodePoint = &HE70E&
        Case "chevronleft": IconCodePoint = &HE76B&
        Case "chevronright": IconCodePoint = &HE76C&
        Case "arrowup": IconCodePoint = &HE74A&
        Case "arrowdown": IconCodePoint = &HE74B&
        Case "back": IconCodePoint = &HE72B&
        Case "forward": IconCodePoint = &HE72A&
        Case "home": IconCodePoint = &HE80F&
        Case "calendar": IconCodePoint = &HE787&
        Case "clock": IconCodePoint = &HE823&
        Case "info": IconCodePoint = &HE946&
        Case "warning": IconCodePoint = &HE7BA&
        Case "error": IconCodePoint = &HE783&
        Case "help": IconCodePoint = &HE897&
        Case "download": IconCodePoint = &HE896&
        Case "upload": IconCodePoint = &HE898&
        Case "share": IconCodePoint = &HE72D&
        Case "send": IconCodePoint = &HE724&
        Case "mail": IconCodePoint = &HE715&
        Case "phone": IconCodePoint = &HE717&
        Case "person": IconCodePoint = &HE77B&
        Case "people": IconCodePoint = &HE716&
        Case "lock": IconCodePoint = &HE72E&
        Case "unlock": IconCodePoint = &HE785&
        Case "star": IconCodePoint = &HE734&
        Case "starfill": IconCodePoint = &HE735&
        Case "heart": IconCodePoint = &HEB51&
        Case "heartfill": IconCodePoint = &HEB52&
        Case "flag": IconCodePoint = &HE7C1&
        Case "tag": IconCodePoint = &HE8EC&
        Case "pin": IconCodePoint = &HE718&
        Case "view": IconCodePoint = &HE890&
        Case "copy": IconCodePoint = &HE8C8&
        Case "paste": IconCodePoint = &HE77F&
        Case "cut": IconCodePoint = &HE8C6&
        Case "undo": IconCodePoint = &HE7A7&
        Case "redo": IconCodePoint = &HE7A6&
        Case "play": IconCodePoint = &HE768&
        Case "pause": IconCodePoint = &HE769&
        Case "stop": IconCodePoint = &HE71A&
        Case "folder": IconCodePoint = &HE8B7&
        Case "document": IconCodePoint = &HE8A5&
        Case "print": IconCodePoint = &HE749&
        Case "attach": IconCodePoint = &HE723&
        Case "link": IconCodePoint = &HE71B&
        Case "globe": IconCodePoint = &HE774&
        Case "cloud": IconCodePoint = &HE753&
        Case "camera": IconCodePoint = &HE722&
        Case "library": IconCodePoint = &HE8F1&
        Case "shop": IconCodePoint = &HE719&
        Case "zoomin": IconCodePoint = &HE8A3&
        Case "zoomout": IconCodePoint = &HE71F&
        Case "list": IconCodePoint = &HE8FD&
        Case "keyboard": IconCodePoint = &HE765&
        Case "power": IconCodePoint = &HE7E8&
    End Select
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
    Set mStateKeys = New Collection
    Set mStateHandlers = New Collection
    Set mDirtyIds = New Collection
    Set mDrawnShapes = New Collection
    mSurfaceFresh = Not SheetHoldsPrefix(hostSheet, SHAPE_PREFIX & targetAppId & "_")
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

' A loading placeholder: a muted block, or with SkeletonLines the bars of
' the text it stands in for, pulsing gently unless motion is reduced.
' Hide it and show the content when the data arrives.
Public Function Skeleton(ByVal targetComponentId As String) As ReDimUI
    Set Skeleton = GetOrCreateComponent(targetComponentId, KIND_SKELETON)
End Function

' A tab strip: one tab per item, the shown tab in bold over an accent
' bar. Controls join a tab's panel with OnTab and show only while their
' tab is the one shown.
Public Function Tabs(ByVal targetComponentId As String) As ReDimUI
    Set Tabs = GetOrCreateComponent(targetComponentId, KIND_TABS)
End Function

' A date field: its face shows the date picked, and a click opens a month
' calendar under it to pick another. WritesTo writes a Date.
Public Function DatePicker(ByVal targetComponentId As String) As ReDimUI
    Set DatePicker = GetOrCreateComponent(targetComponentId, KIND_DATE)
End Function

' A data table: a header row whose cells sort by their column, rows that
' page within the table's height, and one selected row.
Public Function Table(ByVal targetComponentId As String) As ReDimUI
    Set Table = GetOrCreateComponent(targetComponentId, KIND_TABLE)
End Function

' A small pill holding a count or a status word, in the primary color
' unless a variant says otherwise; it widens to its text.
Public Function Badge(ByVal targetComponentId As String) As ReDimUI
    Set Badge = GetOrCreateComponent(targetComponentId, KIND_BADGE)
End Function

' A small trend line through a series of numbers, with a dot on the
' last, in the primary color unless a variant says otherwise.
Public Function Sparkline(ByVal targetComponentId As String) As ReDimUI
    Set Sparkline = GetOrCreateComponent(targetComponentId, KIND_SPARKLINE)
End Function

' A layout container: controls that join it with InStack flow down it in
' the order they joined, or across it with Across, Gap points apart, and
' one that hides gives its place to the next. A caption above a control
' and the note row under a field count in the flow. The stack's height
' follows its content, and so does its width across or when it is 0.
Public Function Stack(ByVal targetComponentId As String) As ReDimUI
    Set Stack = GetOrCreateComponent(targetComponentId, KIND_STACK)
End Function

' A collapsible section: a header whose chevron opens and closes the
' controls that join it with InExpander. Closed, they hide, and in a
' stack the controls after them move up.
Public Function Expander(ByVal targetComponentId As String) As ReDimUI
    Set Expander = GetOrCreateComponent(targetComponentId, KIND_EXPANDER)
End Function

' A button whose click drops a menu of commands, each added with
' AddCommand: a pick runs the command's handler, and the button keeps
' its own text. It shares the drop list's keys, windowing, group
' headers, and disabled items.
Public Function MenuButton(ByVal targetComponentId As String) As ReDimUI
    Set MenuButton = GetOrCreateComponent(targetComponentId, KIND_SELECT)
    MenuButton.MarkMenu
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

' The lowest edge any live toast currently reaches, glides included. -1
' when the tray is otherwise empty.
Friend Function MaxToastVisualBottom(ByVal claimantId As String) As Double
    Dim componentValue As ReDimUI

    MaxToastVisualBottom = -1
    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then
            If componentValue.ComponentId <> claimantId Then
                If componentValue.ToastVisualTop + componentValue.ToastHeight _
                    > MaxToastVisualBottom Then
                    MaxToastVisualBottom = componentValue.ToastVisualTop _
                        + componentValue.ToastHeight
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
' with OK and optional Cancel buttons. No UserForms. The buttons read
' the OK and Cancel UI texts unless given words; "" for cancelText
' leaves Cancel out.
Public Sub Confirm( _
    ByVal titleText As String, _
    ByVal messageText As String, _
    Optional ByVal okProc As String, _
    Optional ByVal cancelProc As String, _
    Optional ByVal okText As Variant, _
    Optional ByVal cancelText As Variant _
)
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim cardLeft As Double
    Dim cardTop As Double
    Dim okWords As String
    Dim cancelWords As String

    RequireRole ROLE_APP, "Confirm"
    If IsMissing(okText) Then
        okWords = Phrase(UI_OK)
    Else
        okWords = CStr(okText)
    End If
    If IsMissing(cancelText) Then
        cancelWords = Phrase(UI_CANCEL)
    Else
        cancelWords = CStr(cancelText)
    End If
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
        .Text(okWords).Primary.Visible True
    Button("mdl_ok").OnClick okProc
    Component("mdl_ok").MarkCloseModalOnClick
    Button("mdl_cancel").AtRect(cardLeft + 340 - 102, cardTop + 150 - 44, 90, 30) _
        .Text(cancelWords).Secondary.Visible(LenB(cancelWords) > 0)
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

' The tab whose panel a Tabs control shows: 0 while the control is
' hidden or has no tabs, so its panels hide with it; -1 when the app
' has no Tabs control by that id, which gates nothing.
Friend Function ShownTabOf(ByVal tabsId As String) As Long
    Dim componentValue As ReDimUI

    RequireRole ROLE_APP, "ShownTabOf"
    ShownTabOf = -1
    If TryGetComponent(tabsId, componentValue) Then
        ShownTabOf = componentValue.PanelTab
    End If
End Function

' Re-applies the controls on a Tabs control's panels that the tab now
' shown hides or shows. One that hides while it has keyboard focus
' commits and lets focus go first.
Friend Sub ReapplyTabMembers(ByVal tabsId As String, ByVal panelNow As Long)
    Dim componentValue As ReDimUI
    Dim shutNow As Boolean

    RequireRole ROLE_APP, "ReapplyTabMembers"
    For Each componentValue In mComponents
        If StrComp(componentValue.TabOwnerId, tabsId, vbTextCompare) = 0 Then
            shutNow = componentValue.TabShutUnder(panelNow)
            If shutNow <> componentValue.TabShut Then
                If shutNow And ReDimUI.IsComponentFocused(mAppId, _
                    componentValue.ComponentId) Then
                    ReDimUI.EndKeyboardFocus
                End If
                componentValue.Apply mSheet, mTheme, Me
            End If
        End If
    Next componentValue
End Sub

' Turns on the command palette: paletteKey, Ctrl+Shift+P unless another
' OnKey code is given ("" for none), opens it while the app's sheet is
' in front, and so does OpenCommandPalette from code.
Public Function CommandPalette(Optional ByVal paletteKey As String = "^+p") As ReDimUI
    RequireRole ROLE_APP, "CommandPalette"
    mPaletteOn = True
    If LenB(paletteKey) > 0 Then HotKey paletteKey, "RdxOpenPalette"
    Set CommandPalette = Me
End Function

Friend Property Get PaletteOn() As Boolean
    PaletteOn = mPaletteOn
End Property

' The command palette: a field at the top of the visible window over a
' list of what the app can do. It lists the app's AddCommand entries,
' then each shown, enabled button with a click handler, then each
' MenuButton command as "Menu: Command"; typing filters them, and a
' pick closes the palette and runs its entry.
Public Sub OpenCommandPalette()
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim cardWidth As Double
    Dim cardLeft As Double
    Dim cardTop As Double
    Dim entryTexts As Collection
    Dim paletteField As ReDimUI

    RequireRole ROLE_APP, "OpenCommandPalette"
    ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    cardWidth = 440
    If cardWidth > viewWidth - 16 Then cardWidth = viewWidth - 16
    If cardWidth < 160 Then cardWidth = 160
    cardLeft = viewLeft + (viewWidth - cardWidth) / 2
    cardTop = viewTop + 24
    Set entryTexts = BuildPaletteEntries()
    mFocusBeforePalette = vbNullString
    If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
        mFocusBeforePalette = ReDimUI.FocusedComponentId
    End If
    BeginUpdate
    Card(PALETTE_CARD_ID).AtRect(cardLeft, cardTop, cardWidth, 48).Text("").Visible True
    Set paletteField = ComboBox(PALETTE_FIELD_ID)
    paletteField.AtRect(cardLeft + 12, cardTop + 12, cardWidth - 24, 24) _
        .Placeholder(Phrase(UI_TYPE_COMMAND)).RestrictToItems.Visible True
    paletteField.ItemsFrom entryTexts
    paletteField.InputValue = vbNullString
    paletteField.OnChange "ReDimHost.RdxPalettePick"
    EndUpdate
    If Not mRendered Then Render
    On Error Resume Next
    mSheet.Shapes(Card(PALETTE_CARD_ID).ShapeName).ZOrder msoBringToFront
    mSheet.Shapes(paletteField.ShapeName).ZOrder msoBringToFront
    On Error GoTo 0
    paletteField.TakeFocus Me, True
End Sub

' What the palette lists, with what each entry runs: the app's commands,
' each shown and enabled button with a click handler, and each menu's
' commands. A text listed once keeps its first entry.
Private Function BuildPaletteEntries() As Collection
    Dim entryTexts As Collection
    Dim componentValue As ReDimUI
    Dim commandWords As Variant

    Set entryTexts = New Collection
    Set mPaletteActions = New Collection
    If Not mCommandOrder Is Nothing Then
        For Each commandWords In mCommandOrder
            AddPaletteEntry entryTexts, CStr(commandWords), "cmd" & vbTab & CStr(commandWords)
        Next commandWords
    End If
    For Each componentValue In mComponents
        componentValue.OfferPaletteEntries Me, entryTexts
    Next componentValue
    Set BuildPaletteEntries = entryTexts
End Function

Friend Sub AddPaletteEntry( _
    ByVal entryTexts As Collection, _
    ByVal entryText As String, _
    ByVal entryAction As String _
)
    Dim existing As Variant

    If LenB(entryText) = 0 Then Exit Sub
    If TryGetItem(mPaletteActions, LCase$(entryText), existing) Then Exit Sub
    mPaletteActions.Add entryAction, LCase$(entryText)
    entryTexts.Add entryText
End Sub

Friend Property Get PaletteOpen() As Boolean
    Dim paletteField As ReDimUI

    If TryGetComponent(PALETTE_FIELD_ID, paletteField) Then
        PaletteOpen = paletteField.IsVisible
    End If
End Property

' The palette's pick: the palette closes, the keys go back where they
' were, and the entry runs.
Friend Sub RunPaletteEntry()
    Dim paletteField As ReDimUI
    Dim found As Variant
    Dim entryParts() As String
    Dim targetValue As ReDimUI
    Dim pickedText As String

    If Not TryGetComponent(PALETTE_FIELD_ID, paletteField) Then Exit Sub
    pickedText = paletteField.InputValue
    If mPaletteActions Is Nothing Then Exit Sub
    If Not TryGetItem(mPaletteActions, LCase$(pickedText), found) Then Exit Sub
    ClosePalette True
    entryParts = Split(CStr(found), vbTab)
    Select Case entryParts(0)
        Case "cmd"
            mLastCommand = entryParts(1)
            If LenB(CommandProcOf(entryParts(1))) > 0 Then
                ReDimUI.RunHandlerAs Me, Nothing, CommandProcOf(entryParts(1))
            End If
        Case "btn"
            If TryGetComponent(entryParts(1), targetValue) Then
                ReDimUI.DispatchShape targetValue.ShapeName
            End If
        Case "menu"
            If TryGetComponent(entryParts(1), targetValue) Then
                targetValue.RunCommandNamed Me, entryParts(2)
            End If
    End Select
End Sub

' Hides the palette; with restoreFocus the control that had the keys
' when it opened takes them back.
Friend Sub ClosePalette(ByVal restoreFocus As Boolean)
    Dim returnTo As ReDimUI

    If ReDimUI.FocusedComponentId = PALETTE_FIELD_ID Then
        If ReDimUI.IsComponentFocused(mAppId, PALETTE_FIELD_ID) Then
            ReDimUI.ClearKeyboardFocus
        End If
    End If
    BeginUpdate
    HideIfExists PALETTE_FIELD_ID
    HideIfExists PALETTE_CARD_ID
    EndUpdate
    If restoreFocus And LenB(mFocusBeforePalette) > 0 Then
        If TryGetComponent(mFocusBeforePalette, returnTo) Then
            If returnTo.IsFocusable Then returnTo.TakeFocus Me, False
        End If
    End If
    mFocusBeforePalette = vbNullString
End Sub

' The palette field is blurring: its own blur finishes first, so the
' palette hides after it.
Friend Sub HidePaletteSoon()
    BeginUpdate
    HideIfExists PALETTE_FIELD_ID
    HideIfExists PALETTE_CARD_ID
    EndUpdate
    mFocusBeforePalette = vbNullString
End Sub

' A stack member changed: the stack lays out again when next asked.
Friend Sub NoteStackTouched(ByVal stackId As String)
    Dim stackValue As ReDimUI

    If TryGetComponent(stackId, stackValue) Then
        stackValue.MarkStackStale
    End If
End Sub

' Whether a stack hides its members: it exists and is hidden.
Friend Function StackHidden(ByVal stackId As String) As Boolean
    Dim stackValue As ReDimUI

    If TryGetComponent(stackId, stackValue) Then
        StackHidden = Not stackValue.IsVisible
    End If
End Function

Friend Sub QueueReflow(ByVal stackId As String)
    Dim existing As Variant

    If mPendingReflows Is Nothing Then Set mPendingReflows = New Collection
    If TryGetItem(mPendingReflows, stackId, existing) Then Exit Sub
    mPendingReflows.Add stackId, stackId
End Sub

' The stacks whose members' flow changed lay out again once the flush has
' drawn everything, each moving only the members whose place changed; a
' stack that grew queues the stack it sits in.
Private Sub ProcessPendingReflows()
    Dim stackId As String
    Dim stackValue As ReDimUI
    Dim roundNo As Long

    If mPendingReflows Is Nothing Then Exit Sub
    Do While mPendingReflows.Count > 0 And roundNo < 100
        roundNo = roundNo + 1
        stackId = CStr(mPendingReflows.Item(1))
        mPendingReflows.Remove 1
        If TryGetComponent(stackId, stackValue) Then
            stackValue.ReflowMembers mSheet, mTheme, Me
        End If
    Loop
End Sub

' A reflow outside a flush, for a member removed from its stack.
Friend Sub ReflowStackNow(ByVal stackId As String)
    Dim stackValue As ReDimUI

    If TryGetComponent(stackId, stackValue) Then
        stackValue.LeaveStack vbNullString
        stackValue.ReflowMembers mSheet, mTheme, Me
    End If
End Sub

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
    Dim known As Variant

    RequireRole ROLE_APP, "SetState"
    If Not TryGetItem(mStateValues, stateKey, known) Then mStateKeys.Add stateKey
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

' Every state key in the order it was first set, as an array from 0;
' an empty store gives an empty array. With State, enough to save the
' whole store wherever the host keeps it.
Public Property Get StateKeys() As Variant
    Dim keyList() As String
    Dim keyAt As Long

    RequireRole ROLE_APP, "StateKeys"
    If mStateKeys.Count = 0 Then
        StateKeys = Array()
        Exit Property
    End If
    ReDim keyList(0 To mStateKeys.Count - 1)
    For keyAt = 1 To mStateKeys.Count
        keyList(keyAt - 1) = mStateKeys.Item(keyAt)
    Next keyAt
    StateKeys = keyList
End Property

' The state store is deliberately in-memory only. Durability is the host
' application's concern: walk the store with StateKeys and State and save
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
    ' Several controls, or one list or calendar opening or closing (a dozen
    ' parts drawn or deleted at once), paint once at the end rather than
    ' part by part. A single control's usual repaint, such as a keystroke's,
    ' keeps the screen on, since the toggle redraws the scroll bars.
    If Not ReDimUI.InPumpTick Then
        If mDirtyIds.Count > 1 Then
            toggledScreen = True
        ElseIf TryGetComponent(CStr(mDirtyIds.Item(1)), componentValue) Then
            toggledScreen = componentValue.DrawsManyParts
        End If
        If toggledScreen Then Application.ScreenUpdating = False
    End If
    On Error GoTo RestoreScreen
    For Each dirtyId In mDirtyIds
        If TryGetComponent(CStr(dirtyId), componentValue) Then
            componentValue.Apply mSheet, mTheme, Me
        End If
    Next dirtyId
    Set mDirtyIds = New Collection
    ProcessPendingReflows
    If mToastRestackPending Then
        mToastRestackPending = False
        CompactToastSlots
    End If
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

' Hover and press looks for the app's controls, and an open list whose
' highlight follows the pointer. Off by default: while on, the pump reads
' the pointer every frame the app's sheet is in front.
Public Function PointerEffects(Optional ByVal effectsOn As Boolean = True) As ReDimUI
    RequireRole ROLE_APP, "PointerEffects"
    mPointerEffects = effectsOn
    If Not effectsOn Then ClearPointerTarget
    If mRendered And HasAppWork() Then ReDimUI.RequestPump
    Set PointerEffects = Me
End Function

' Follows the Windows look: the app takes ThemeSystem now, and again
' when Windows switches between light and dark mode or changes its
' accent, which the app notices as its sheets activate, the selection
' moves, or the pump runs, at most every two seconds.
Public Function FollowSystemTheme(Optional ByVal followOn As Boolean = True) As ReDimUI
    RequireRole ROLE_APP, "FollowSystemTheme"
    mFollowSystem = followOn
    mAppliedSystemLook = vbNullString
    If followOn Then ReDimUI.CheckSystemLook True
    Set FollowSystemTheme = Me
End Function

Friend Property Get FollowsSystemTheme() As Boolean
    FollowsSystemTheme = mFollowSystem
End Property

' Takes the Windows look when it differs from the one the app last took.
' Before the first Render the theme is only recorded, so following the
' look during a build does not render early.
Friend Sub MatchSystemLook(ByVal lookKey As String)
    If lookKey = mAppliedSystemLook Then Exit Sub
    mAppliedSystemLook = lookKey
    If mRendered Then
        SetTheme ReDimUI.ThemeSystem
        Exit Sub
    End If
    Set mTheme = ReDimUI.ThemeSystem
    If mCanvasPainted Then
        On Error Resume Next
        mSheet.Cells.Interior.Color = mTheme.CanvasColor
        On Error GoTo 0
    End If
End Sub

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

' Drawing-object protection refuses Shape.Duplicate even under
' UserInterfaceOnly. A surface ReDim protected lifts for a copy's moment,
' about a tenth of a millisecond with the protection put back as
' ProtectSurface set it (the selection rule survives). True when lifted.
Friend Function LiftSurfaceProtection() As Boolean
    If Not mSurfaceProtected Then Exit Function
    On Error Resume Next
    mSheet.Unprotect
    LiftSurfaceProtection = (Err.Number = 0)
    On Error GoTo 0
End Function

Friend Sub RestoreSurfaceProtection()
    mSheet.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
End Sub

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
    ReDimUI.ReadButtons leftButtonDown, anyButtonDown
    pressEdge = (anyButtonDown And Not mFieldPrevButtonDown)
    mFieldPrevButtonDown = anyButtonDown
    If ReDimUI.ConsumeForcedPressEdge() Then pressEdge = True
    Set expiredIds = New Collection
    For Each componentValue In Tickables()
        componentValue.TickAnimation mSheet, elapsedMs
        componentValue.TickPressDrag Me, leftButtonDown
        componentValue.TickListWatch pressEdge, runWork, anyButtonDown
        If runWork Then
            If componentValue.IsExpiredToast(nowMs) Then
                componentValue.BeginToastExit
            End If
        End If
        If componentValue.IsToastExitDone Then
            expiredIds.Add componentValue.ComponentId
        End If
    Next componentValue
    TickPointer nowMs, leftButtonDown
    TickHold nowMs, leftButtonDown
    ' The control holding keyboard focus, tickable or not, watches for a
    ' press elsewhere or a moved selection.
    If ReDimUI.IsComponentFocused(mAppId, ReDimUI.FocusedComponentId) Then
        If TryGetComponent(ReDimUI.FocusedComponentId, componentValue) Then
            componentValue.TickFocusWatch Me, pressEdge, runWork, anyButtonDown
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

' True while the app reads the pointer every frame: pointer effects on
' or a control with a tooltip, with its sheet in front.
Private Function NeedsPointerWatch() As Boolean
    If Not (mPointerEffects Or mHasTooltips) Then Exit Function
    NeedsPointerWatch = AppSheetInFront()
End Function

' The control under the pointer takes the hover look, and the press look
' while the left button that went down on it stays down; an open list
' under a moving pointer moves its highlight to the row there.
Private Sub TickPointer(ByVal nowMs As LongLong, ByVal leftButtonDown As Boolean)
    Dim pointsX As Double
    Dim pointsY As Double
    Dim hitId As String
    Dim hitPart As String
    Dim pressedNow As Boolean
    Dim pointerMoved As Boolean
    Dim componentValue As ReDimUI

    If Not NeedsPointerWatch() Then
        If LenB(mHoverTargetId) > 0 Then ClearPointerTarget
        If LenB(mTipShownId) > 0 Then HideTooltip
        mTipTargetId = vbNullString
        Exit Sub
    End If
    If ReDimUI.FramePointer(pointsX, pointsY) Then
        HitTestPoint pointsX, pointsY, hitId, hitPart
        pointerMoved = (pointsX <> mHoverPointX Or pointsY <> mHoverPointY)
        mHoverPointX = pointsX
        mHoverPointY = pointsY
    End If
    If leftButtonDown And Not mHoverPrevDown Then
        mPressTargetId = hitId
        mPressTargetPart = hitPart
    ElseIf Not leftButtonDown Then
        mPressTargetId = vbNullString
        mPressTargetPart = vbNullString
    End If
    mHoverPrevDown = leftButtonDown
    pressedNow = leftButtonDown And LenB(hitId) > 0 _
        And hitId = mPressTargetId And hitPart = mPressTargetPart
    If hitId <> mHoverTargetId Or hitPart <> mHoverTargetPart Then
        ClearPointerTarget
        mHoverTargetId = hitId
        mHoverTargetPart = hitPart
        mHoverPressedShown = pressedNow
        If TryGetComponent(hitId, componentValue) Then
            componentValue.SetPointerLook True, hitPart, pressedNow
        End If
    ElseIf pressedNow <> mHoverPressedShown Then
        mHoverPressedShown = pressedNow
        If TryGetComponent(hitId, componentValue) Then
            componentValue.SetPointerLook True, hitPart, pressedNow
        End If
    End If
    If pointerMoved And LenB(hitId) > 0 Then
        If TryGetComponent(hitId, componentValue) Then
            componentValue.FollowListPointer hitPart
        End If
    End If
    TickTooltip nowMs, hitId, leftButtonDown, pointsX, pointsY
End Sub

' A tooltip shows once the pointer has rested on its control for the
' Windows tooltip delay, the double-click time; a move of more than a few
' points starts the wait over, so a pointer passing through raises none.
' It goes when the pointer leaves the control, a button goes down, the
' control opens its list, or the pointer reaches the tip itself, so a
' click there lands on what the tip covered. After any of those it stays
' away until the pointer leaves the control.
Private Sub TickTooltip( _
    ByVal nowMs As LongLong, _
    ByVal hitId As String, _
    ByVal leftButtonDown As Boolean, _
    ByVal pointsX As Double, _
    ByVal pointsY As Double _
)
    Dim componentValue As ReDimUI

    If hitId <> mTipTargetId Then
        HideTooltip
        mTipTargetId = hitId
        mTipSinceMs = nowMs
        mTipRestX = pointsX
        mTipRestY = pointsY
        mTipSuppressed = False
    End If
    If leftButtonDown Then
        HideTooltip
        mTipSuppressed = True
        Exit Sub
    End If
    If mTipSuppressed Or LenB(hitId) = 0 Then Exit Sub
    If LenB(mTipShownId) > 0 Then
        If pointsX >= mTipLeft - 2 And pointsX <= mTipRight + 2 _
            And pointsY >= mTipTop - 2 And pointsY <= mTipBottom + 2 Then
            HideTooltip
            mTipSuppressed = True
        ElseIf TryGetComponent(mTipShownId, componentValue) Then
            If componentValue.ListIsOpen Then
                HideTooltip
                mTipSuppressed = True
            End If
        End If
        Exit Sub
    End If
    If Abs(pointsX - mTipRestX) > TIP_REST_SLOP Or Abs(pointsY - mTipRestY) > TIP_REST_SLOP Then
        mTipRestX = pointsX
        mTipRestY = pointsY
        mTipSinceMs = nowMs
        Exit Sub
    End If
    If nowMs - mTipSinceMs < GetDoubleClickTime() Then Exit Sub
    If Not TryGetComponent(hitId, componentValue) Then Exit Sub
    If componentValue.DrawTooltip(pointsX, pointsY, mTipLeft, mTipTop, mTipRight, _
        mTipBottom) Then mTipShownId = hitId
End Sub

Private Sub HideTooltip()
    Dim componentValue As ReDimUI

    If LenB(mTipShownId) = 0 Then Exit Sub
    If TryGetComponent(mTipShownId, componentValue) Then
        componentValue.ClearTooltipPart
    End If
    mTipShownId = vbNullString
End Sub

' A click on a tooltip puts it away, as a press anywhere does.
Friend Sub DismissTooltip()
    HideTooltip
    mTipSuppressed = True
End Sub

' A control got a tooltip: the pointer watch starts, and a rendered app
' arms the pump for it.
Friend Sub NoteTooltips()
    mHasTooltips = True
    If mRendered And HasAppWork() Then ReDimUI.RequestPump
End Sub

' Hold-to-repeat: a press held on a stepper's minus or plus, a drop
' list's pager, or a transfer panel's paging arrow repeats its action
' after the Windows keyboard repeat delay, at the keyboard repeat rate,
' while the pointer stays on the part. The release then fires OnChange
' once and swallows the click Excel delivers for it.
Private Sub TickHold(ByVal nowMs As LongLong, ByVal leftButtonDown As Boolean)
    Dim pointsX As Double
    Dim pointsY As Double
    Dim hitId As String
    Dim hitPart As String
    Dim componentValue As ReDimUI

    If leftButtonDown And Not mHoldPrevDown Then
        mHoldId = vbNullString
        If AppSheetInFront() Then
            If ReDimUI.FramePointer(pointsX, pointsY) Then
                HitTestPoint pointsX, pointsY, hitId, hitPart
            End If
            If TryGetComponent(hitId, componentValue) Then
                If componentValue.RepeatsPart(hitPart) Then
                    mHoldId = hitId
                    mHoldPart = hitPart
                    mHoldNextMs = nowMs + HoldDelayMs()
                    mHoldRepeats = 0
                    componentValue.BeginHold
                End If
            End If
        End If
    ElseIf leftButtonDown And LenB(mHoldId) > 0 Then
        If nowMs >= mHoldNextMs Then
            If ReDimUI.FramePointer(pointsX, pointsY) Then
                HitTestPoint pointsX, pointsY, hitId, hitPart
            End If
            If hitId = mHoldId And hitPart = mHoldPart Then
                If TryGetComponent(mHoldId, componentValue) Then
                    componentValue.RepeatPart Me, mHoldPart
                    mHoldRepeats = mHoldRepeats + 1
                End If
            End If
            mHoldNextMs = nowMs + HoldRepeatMs()
        End If
    ElseIf Not leftButtonDown And LenB(mHoldId) > 0 Then
        If TryGetComponent(mHoldId, componentValue) Then
            componentValue.EndHold Me, mHoldRepeats > 0
        End If
        mHoldId = vbNullString
    End If
    mHoldPrevDown = leftButtonDown
End Sub

' The Windows keyboard repeat delay, 250 to 1000 ms, and the interval
' of its repeat rate, from about 400 ms down to 33 ms.
Private Function HoldDelayMs() As Long
    Dim delayCode As Long

    delayCode = 1
    If SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, delayCode, 0) = 0 Then
        delayCode = 1
    End If
    If delayCode < 0 Or delayCode > 3 Then delayCode = 1
    HoldDelayMs = (delayCode + 1) * 250
End Function

Private Function HoldRepeatMs() As Long
    Dim speedCode As Long

    speedCode = 31
    If SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, speedCode, 0) = 0 Then
        speedCode = 31
    End If
    If speedCode < 0 Or speedCode > 31 Then speedCode = 31
    HoldRepeatMs = CLng(1000 / (2.5 + speedCode * 27.5 / 31))
End Function

' The control under the pointer loses its look, as when the pointer
' leaves it or the watch stops.
Private Sub ClearPointerTarget()
    Dim componentValue As ReDimUI

    If LenB(mHoverTargetId) > 0 Then
        If TryGetComponent(mHoverTargetId, componentValue) Then
            componentValue.SetPointerLook False, vbNullString, False
        End If
    End If
    mHoverTargetId = vbNullString
    mHoverTargetPart = vbNullString
    mHoverPressedShown = False
End Sub

' The control under a sheet point and its part there. An open list
' draws over everything, so it answers first; then the last control
' created that covers the point, since later shapes draw on top.
Private Sub HitTestPoint( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double, _
    ByRef hitId As String, _
    ByRef hitPart As String _
)
    Dim componentValue As ReDimUI
    Dim topValue As ReDimUI

    For Each componentValue In mComponents
        If componentValue.ListHoldsPoint(pointsX, pointsY) Then
            Set topValue = componentValue
            Exit For
        End If
    Next componentValue
    If topValue Is Nothing Then
        For Each componentValue In mComponents
            If componentValue.CoversPoint(pointsX, pointsY) Then
                Set topValue = componentValue
            End If
        Next componentValue
    End If
    If topValue Is Nothing Then Exit Sub
    hitId = topValue.ComponentId
    hitPart = topValue.PartAtPoint(pointsX, pointsY)
End Sub

' After any toast leaves, survivors renumber to slots 1..N in their current
' order and slide up into the freed positions, the way a notification tray
' behaves. Each sits under the toasts above it, heights and gaps summed,
' since a toast is as tall as its message; the same pass restacks the
' tray when a toast's height changes.
Private Sub CompactToastSlots()
    Dim live As Collection
    Dim componentValue As ReDimUI
    Dim lowest As ReDimUI
    Dim idx As Long
    Dim lowestAt As Long
    Dim assigned As Long
    Dim stackAbove As Double

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
        lowest.ReslotToast assigned, stackAbove
        stackAbove = stackAbove + lowest.ToastHeight + TOAST_GAP
        live.Remove lowestAt
    Loop
End Sub

' A toast's height changed while drawing; the tray restacks after the
' flush, not in the middle of it.
Friend Sub NoteToastResized()
    mToastRestackPending = True
End Sub

' How far down the tray a slot sits: the live toasts in the slots above
' it, each with its height and gap.
Friend Function ToastStackAbove(ByVal claimantId As String, ByVal slot As Long) As Double
    Dim componentValue As ReDimUI

    For Each componentValue In mComponents
        If componentValue.IsPendingToast Then
            If componentValue.ComponentId <> claimantId Then
                If componentValue.ToastSlot < slot Then
                    ToastStackAbove = ToastStackAbove + componentValue.ToastHeight + TOAST_GAP
                End If
            End If
        End If
    Next componentValue
End Function

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
    If NeedsPointerWatch() Then
        HasAppWork = True
        Exit Function
    End If
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
        If componentValue.NeedsPressWatch Then
            If AppSheetInFront() Then
                HasAppWork = True
                Exit Function
            End If
        End If
    Next componentValue
End Function

Private Function AppSheetInFront() As Boolean
    On Error Resume Next
    AppSheetInFront = (mSheet Is ActiveSheet)
    On Error GoTo 0
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
        Case KIND_SKELETON
            mWidth = 200: mHeight = 60
        Case KIND_DATE
            mWidth = 130: mHeight = 24
        Case KIND_TABLE
            mWidth = 420: mHeight = 200
        Case KIND_BADGE
            mWidth = 0: mHeight = 18
        Case KIND_SPARKLINE
            mWidth = 120: mHeight = 32
        Case KIND_STACK
            mWidth = 300: mHeight = 0
            mStackGap = 8
            mStackLaidGen = -1
            mAppliedStackShown = True
        Case KIND_EXPANDER
            mWidth = 300: mHeight = 28
            mFontBold = True
            mAppliedPanelTab = -2
        Case KIND_TABS
            mWidth = 360: mHeight = 32
            mValue = 1
            mAppliedPanelTab = -2
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
            KIND_SELECT, KIND_STEPPER, KIND_TRANSFER, KIND_SKELETON
            IsTickable = True
        Case KIND_TOGGLE, KIND_TABS
            ' Only while a knob or a tab bar glides.
            IsTickable = (LenB(mGlidePart) > 0)
        Case KIND_DATE
            ' Only while the calendar is open, for its click-away watch.
            IsTickable = mSelectOpen
    End Select
End Property

Public Property Get OwnerApp() As ReDimUI
    RequireRole ROLE_COMPONENT, "OwnerApp"
    Set OwnerApp = ReDimUI.App(mOwnerAppId)
End Property

Public Function At(ByVal rangeAddress As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "At"
    mAnchor = rangeAddress
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
    If mKind = KIND_STACK Then
        mStackAutoWidth = (widthPoints <= 0)
        MarkStackStale
    End If
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
    If mKind = KIND_STACK Then
        mStackAutoWidth = (widthPoints <= 0)
        MarkStackStale
    End If
    Touch
    Set Sized = Me
End Function

Public Function Text(ByVal displayText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Text"
    mText = displayText
    Touch
    Set Text = Me
End Function

' What a busy button shows, the Working UI text unless given; ""
' keeps the button's own text.
Public Function BusyText(ByVal displayText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BusyText"
    mBusyText = displayText
    mBusyTextGiven = True
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
    ' While its tab hides the control, the flag waits for the tab.
    If mTabShut Then
        mVisibleBeforeTab = flag
    Else
        mVisible = flag
    End If
    Touch
    Set Visible = Me
End Function

Public Function Enabled(ByVal flag As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "Enabled"
    mEnabled = flag
    Touch
    Set Enabled = Me
End Function

' Puts the control on a Tabs control's panel: it shows only while tab
' tabNumber is the one shown, and hides with the Tabs control. Visible
' still decides whether it shows on its tab; OnTab "" takes it off.
Public Function OnTab(ByVal tabsId As String, ByVal tabNumber As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "OnTab"
    mTabOwnerId = tabsId
    mTabNumber = tabNumber
    Touch
    Set OnTab = Me
End Function

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

' The border of a toast's action button, an RGB color; the theme's accent
' until set.
Public Function ActionBorder(ByVal borderRgb As Long) As ReDimUI
    RequireToastKind "ActionBorder"
    mToastActionBorder = borderRgb
    Touch
    Set ActionBorder = Me
End Function

' The narrowest and widest a toast may be, in points. A toast fits its
' message between the two and wraps past the widest; both are 240 until
' set, a fixed width. A MinWidth under the MaxWidth lets a short message
' make a narrow toast, and a MinWidth over it fixes the wider width.
Public Function MinWidth(ByVal widthPoints As Double) As ReDimUI
    RequireToastKind "MinWidth"
    mToastMinWidth = widthPoints
    Touch
    Set MinWidth = Me
End Function

Public Function MaxWidth(ByVal widthPoints As Double) As ReDimUI
    RequireToastKind "MaxWidth"
    mToastMaxWidth = widthPoints
    Touch
    Set MaxWidth = Me
End Function

' The tallest a toast grows to fit a wrapped message, in points, 160
' until set; a longer message ends in an ellipsis. Never less than one
' line's 40 points.
Public Function MaxHeight(ByVal heightPoints As Double) As ReDimUI
    RequireToastKind "MaxHeight"
    mToastMaxHeight = heightPoints
    Touch
    Set MaxHeight = Me
End Function

Private Sub RequireToastKind(ByVal memberName As String)
    RequireRole ROLE_COMPONENT, memberName
    If mKind <> KIND_TOAST Then
        RaiseError ERROR_INVALID_ARGUMENT, memberName, memberName & " applies to toasts."
    End If
End Sub

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

' A keyboard shortcut that clicks the control while its sheet is in
' front, as an Application.OnKey code: "^s" for Ctrl+S, "^+e" for
' Ctrl+Shift+E, "{F5}" for F5. It needs Ctrl or Alt, or is a function key,
' so it never takes a key a focused field types; "" removes it. The
' control's tooltip and alternative text name it.
Public Function Shortcut(ByVal keyCode As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Shortcut"
    If LenB(keyCode) > 0 And Not ShortcutCodeValid(keyCode) Then
        RaiseError ERROR_INVALID_ARGUMENT, "Shortcut", "Shortcut takes an Application.OnKey " & _
            "code with Ctrl or Alt, such as ""^s"", or a function key, such as ""{F5}""."
    End If
    mShortcut = keyCode
    If LenB(keyCode) > 0 Then OwnerApp.NoteTooltips
    Touch
    ReDimUI.RefreshAccessKeys
    Set Shortcut = Me
End Function

Friend Property Get ShortcutCode() As String
    ShortcutCode = mShortcut
End Property

' Ctrl or Alt with one key or a braced key name, or a bare function key.
' A quote or an apostrophe would break the OnKey procedure string.
Private Function ShortcutCodeValid(ByVal keyCode As String) As Boolean
    Dim pos As Long
    Dim modifiers As String
    Dim keyPart As String

    If InStr(keyCode, """") > 0 Or InStr(keyCode, "'") > 0 Then Exit Function
    pos = 1
    Do While pos <= Len(keyCode)
        If InStr("^%+", Mid$(keyCode, pos, 1)) = 0 Then Exit Do
        modifiers = modifiers & Mid$(keyCode, pos, 1)
        pos = pos + 1
    Loop
    keyPart = Mid$(keyCode, pos)
    If LenB(keyPart) = 0 Then Exit Function
    If Len(keyPart) > 1 And Not (Left$(keyPart, 1) = "{" And Right$(keyPart, 1) = "}") Then
        Exit Function
    End If
    If InStr(modifiers, "^") > 0 Or InStr(modifiers, "%") > 0 Then
        ShortcutCodeValid = True
    ElseIf UCase$(keyPart) Like "{F#}" Or UCase$(keyPart) Like "{F##}" Then
        ShortcutCodeValid = True
    End If
End Function

' A shortcut's OnKey code as people read it: "^+s" is Ctrl+Shift+S, "{F5}"
' is F5, "%{DEL}" is Alt+Del.
Private Function ShortcutLabel(ByVal keyCode As String) As String
    Dim pos As Long
    Dim modifiers As String
    Dim keyPart As String

    pos = 1
    Do While pos <= Len(keyCode)
        If InStr("^%+", Mid$(keyCode, pos, 1)) = 0 Then Exit Do
        modifiers = modifiers & Mid$(keyCode, pos, 1)
        pos = pos + 1
    Loop
    keyPart = Mid$(keyCode, pos)
    If Len(keyPart) > 2 And Left$(keyPart, 1) = "{" And Right$(keyPart, 1) = "}" Then
        keyPart = Mid$(keyPart, 2, Len(keyPart) - 2)
        Select Case UCase$(keyPart)
            Case "DEL", "DELETE"
                keyPart = "Del"
            Case "BS", "BACKSPACE"
                keyPart = "Backspace"
            Case "ESC", "ESCAPE"
                keyPart = "Esc"
            Case "INSERT", "INS"
                keyPart = "Ins"
            Case "PGUP"
                keyPart = "PgUp"
            Case "PGDN"
                keyPart = "PgDn"
            Case Else
                If UCase$(keyPart) Like "F#" Or UCase$(keyPart) Like "F##" Or Len(keyPart) = 1 Then
                    keyPart = UCase$(keyPart)
                Else
                    keyPart = UCase$(Left$(keyPart, 1)) & LCase$(Mid$(keyPart, 2))
                End If
        End Select
    ElseIf keyPart = "~" Then
        keyPart = "Enter"
    Else
        keyPart = UCase$(keyPart)
    End If
    If InStr(modifiers, "^") > 0 Then ShortcutLabel = "Ctrl+"
    If InStr(modifiers, "%") > 0 Then ShortcutLabel = ShortcutLabel & "Alt+"
    If InStr(modifiers, "+") > 0 Then ShortcutLabel = ShortcutLabel & "Shift+"
    ShortcutLabel = ShortcutLabel & keyPart
End Function

' Shows a float TextInput's characters as dots, as a password box does.
' Copying and cutting take nothing; the text still reads through
' InputValue and writes to WritesTo.
Public Function Masked(Optional ByVal maskedOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Masked"
    mMasked = maskedOn
    Touch
    Set Masked = Me
End Function

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

' A caption above the control, in small text; leave room for it above
' the control's rectangle.
Public Function Caption(ByVal captionWords As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Caption"
    mCaptionText = captionWords
    Touch
    Set Caption = Me
End Function

' Helper text below the control in muted ink. A float field's validation
' or error message takes its place while one shows.
Public Function Hint(ByVal hintText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Hint"
    mHintText = hintText
    Touch
    Set Hint = Me
End Function

' Marks the control required: an asterisk after its caption, in the
' danger color. A float TextInput or ComboBox committed empty shows
' emptyMessage as its validation message, the Required UI text unless
' given.
Public Function Required( _
    Optional ByVal requiredOn As Boolean = True, _
    Optional ByVal emptyMessage As String _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "Required"
    mRequired = requiredOn
    mRequiredMessage = emptyMessage
    Touch
    Set Required = Me
End Function

' An error from outside the field, a server's answer for one: a float
' TextInput or ComboBox shows it with a danger border until it is set
' to "". A validation message shows in its place while there is one.
Public Function ErrorText(ByVal messageText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "ErrorText"
    mErrorText = messageText
    Touch
    Set ErrorText = Me
End Function

' The text lines a Skeleton stands in for, drawn as bars with the last
' one shorter; 0 draws one block.
Public Function SkeletonLines(ByVal lineCount As Long) As ReDimUI
    RequireRole ROLE_COMPONENT, "SkeletonLines"
    If mKind <> KIND_SKELETON Then
        RaiseError ERROR_INVALID_OPERATION, "SkeletonLines", _
            "SkeletonLines applies to Skeleton only."
    End If
    If lineCount < 0 Then lineCount = 0
    mSkeletonLines = lineCount
    Touch
    Set SkeletonLines = Me
End Function

' Sets the date a DatePicker shows, as a program change: nothing is
' written to WritesTo and OnChange does not fire. Any time of day is
' dropped; Value 0 clears the date.
Public Function PickDate(ByVal chosenDate As Date) As ReDimUI
    RequireDatePicker "PickDate"
    mValue = CDbl(DayOnly(chosenDate))
    Touch
    Set PickDate = Me
End Function

' The date a DatePicker holds, as a Date, or Empty while none is picked.
Public Property Get PickedDate() As Variant
    RequireDatePicker "PickedDate"
    If mValue <> 0 Then PickedDate = CDate(mValue)
End Property

' The earliest and latest days the calendar picks; days outside read
' muted and take no clicks. A side left out stays open.
Public Function DateRange( _
    Optional ByVal earliest As Variant, _
    Optional ByVal latest As Variant _
) As ReDimUI
    RequireDatePicker "DateRange"
    mDateEarliest = 0
    mDateLatest = 0
    If Not IsMissing(earliest) Then mDateEarliest = DayOnly(CDate(earliest))
    If Not IsMissing(latest) Then mDateLatest = DayOnly(CDate(latest))
    Touch
    Set DateRange = Me
End Function

' The face's Format$ pattern, such as "yyyy-mm-dd" or "d mmm yyyy";
' the system's short date by default, and "" goes back to it.
Public Function DateFormat(ByVal formatPattern As String) As ReDimUI
    RequireDatePicker "DateFormat"
    mDateFormat = formatPattern
    Touch
    Set DateFormat = Me
End Function

' The header of each column, which also sets how many columns the table
' has; rows already held keep their cells in the columns that remain.
Public Function Columns(ParamArray heads() As Variant) As ReDimUI
    Dim colNo As Long

    RequireTable "Columns"
    ResizeTableColumns UBound(heads) - LBound(heads) + 1
    For colNo = 1 To mTableCols
        mTableHeads(colNo) = CStr(heads(LBound(heads) + colNo - 1))
    Next colNo
    TableChanged
    Set Columns = Me
End Function

' Appends a row, one value per column in order; missing values leave
' their cells empty and extra ones are dropped. A table with no columns
' yet takes one per value.
Public Function AddRow(ParamArray rowCells() As Variant) As ReDimUI
    Dim colNo As Long
    Dim valueTotal As Long

    RequireTable "AddRow"
    valueTotal = UBound(rowCells) - LBound(rowCells) + 1
    If mTableCols = 0 Then ResizeTableColumns valueTotal
    If mTableCols = 0 Then
        Set AddRow = Me
        Exit Function
    End If
    EnsureTableRoom mTableRows + 1
    mTableRows = mTableRows + 1
    For colNo = 1 To mTableCols
        If colNo <= valueTotal Then
            mTableCells(colNo, mTableRows) = rowCells(LBound(rowCells) + colNo - 1)
        Else
            mTableCells(colNo, mTableRows) = Empty
        End If
    Next colNo
    TableChanged
    Set AddRow = Me
End Function

' Replaces the rows from a Range or a two-dimensional array. With
' firstRowHeads the first row names the columns; without it, a table
' that already has as many columns keeps their headers.
Public Function TableFrom( _
    ByVal tableSource As Variant, _
    Optional ByVal firstRowHeads As Boolean = True _
) As ReDimUI
    Dim grid As Variant
    Dim rowLow As Long
    Dim colLow As Long
    Dim sourceRows As Long
    Dim sourceCols As Long
    Dim rowNo As Long
    Dim colNo As Long
    Dim skipRows As Long
    Dim rankFailed As Boolean

    RequireTable "TableFrom"
    If IsObject(tableSource) Then
        If Not TypeOf tableSource Is Range Then rankFailed = True
    ElseIf Not IsArrayValue(tableSource) Then
        rankFailed = True
    End If
    If Not rankFailed Then
        If IsObject(tableSource) Then
            If tableSource.Cells.CountLarge = 1 Then
                ReDim grid(1 To 1, 1 To 1)
                grid(1, 1) = tableSource.Value
            Else
                grid = tableSource.Value
            End If
        Else
            grid = tableSource
        End If
        On Error Resume Next
        colLow = LBound(grid, 2)
        rankFailed = (Err.Number <> 0)
        On Error GoTo 0
    End If
    If rankFailed Then
        RaiseError ERROR_INVALID_ARGUMENT, "TableFrom", _
            "TableFrom takes a Range or a two-dimensional array."
    End If
    rowLow = LBound(grid, 1)
    sourceRows = UBound(grid, 1) - rowLow + 1
    sourceCols = UBound(grid, 2) - colLow + 1
    If firstRowHeads Then skipRows = 1
    If firstRowHeads Or sourceCols <> mTableCols Then
        ResizeTableColumns sourceCols
        For colNo = 1 To sourceCols
            If firstRowHeads Then
                mTableHeads(colNo) = TableCellText(0, grid(rowLow, colLow + colNo - 1))
            Else
                mTableHeads(colNo) = Phrase(UI_COLUMN, colNo)
            End If
        Next colNo
    End If
    mTableRows = 0
    If sourceRows - skipRows > 0 Then
        EnsureTableRoom sourceRows - skipRows
        For rowNo = 1 + skipRows To sourceRows
            mTableRows = mTableRows + 1
            For colNo = 1 To sourceCols
                mTableCells(colNo, mTableRows) = grid(rowLow + rowNo - 1, colLow + colNo - 1)
            Next colNo
        Next rowNo
    End If
    mValue = 0
    mTableOffset = 0
    TableChanged
    Set TableFrom = Me
End Function

' Drops every row and the selection; the columns stay.
Public Function ClearRows() As ReDimUI
    RequireTable "ClearRows"
    mTableRows = 0
    mValue = 0
    mTableOffset = 0
    TableChanged
    Set ClearRows = Me
End Function

' Column widths in points, in column order; 0 or a column left out
' shares the width left over by the length of its content.
Public Function ColumnWidths(ParamArray columnPoints() As Variant) As ReDimUI
    Dim colNo As Long

    RequireTable "ColumnWidths"
    For colNo = 1 To mTableCols
        mColumnPoints(colNo) = 0
        If colNo <= UBound(columnPoints) - LBound(columnPoints) + 1 Then
            mColumnPoints(colNo) = CDbl(columnPoints(LBound(columnPoints) + colNo - 1))
        End If
    Next colNo
    mColumnLayoutKey = vbNullString
    Touch
    Set ColumnWidths = Me
End Function

' How a column shows its values, as a Format$ pattern such as "#,##0.00"
' or "yyyy-mm-dd"; sorting still compares the values themselves.
Public Function ColumnFormat( _
    ByVal columnNumber As Long, _
    ByVal formatPattern As String _
) As ReDimUI
    RequireTable "ColumnFormat"
    If columnNumber >= 1 And columnNumber <= mTableCols Then
        mColumnFormats(columnNumber) = formatPattern
        mColumnLayoutKey = vbNullString
        Touch
    End If
    Set ColumnFormat = Me
End Function

' Sorts the rows by a column, as a click on its header does, without
' firing anything; 0 goes back to the order the rows came in.
Public Function SortBy( _
    ByVal columnNumber As Long, _
    Optional ByVal sortDescending As Boolean = False _
) As ReDimUI
    RequireTable "SortBy"
    If columnNumber < 0 Or columnNumber > mTableCols Then columnNumber = 0
    mSortColumn = columnNumber
    mSortDescending = sortDescending And columnNumber > 0
    mTableOrderCurrent = False
    mTableOffset = 0
    Touch
    Set SortBy = Me
End Function

' The words the empty row shows while the table holds no rows.
Public Function EmptyText(ByVal emptyWords As String) As ReDimUI
    RequireTable "EmptyText"
    mTableEmptyText = emptyWords
    Touch
    Set EmptyText = Me
End Function

' The Public procedure a double click on a row, or Enter on the
' selected row while the table has the keys, runs, with the table as
' ReDimUI.Sender and the row in its CurrentValue.
Public Function OnRowOpen(ByVal handlerProc As String) As ReDimUI
    RequireTable "OnRowOpen"
    mRowOpenProc = handlerProc
    Set OnRowOpen = Me
End Function

' Filters the rows from code, as typing does while the table has the
' keys: a row shows when any of its cells, as shown, holds the text,
' case aside. "" shows every row.
Public Function FilterRows(ByVal filterText As String) As ReDimUI
    RequireTable "FilterRows"
    SetTableFilter filterText
    Set FilterRows = Me
End Function

' Writes the header and the rows the table shows, filtered and sorted,
' to the sheet from exportAnchor's top-left cell down, values as they
' are, in one write.
Public Function ExportTo(ByVal exportAnchor As Range) As ReDimUI
    Dim grid() As Variant
    Dim rowNo As Long
    Dim colNo As Long

    RequireTable "ExportTo"
    If mTableCols = 0 Then
        Set ExportTo = Me
        Exit Function
    End If
    EnsureTableShown
    ReDim grid(1 To mTableShownCount + 1, 1 To mTableCols)
    For colNo = 1 To mTableCols
        grid(1, colNo) = mTableHeads(colNo)
        For rowNo = 1 To mTableShownCount
            grid(rowNo + 1, colNo) = mTableCells(colNo, mTableShown(rowNo))
        Next rowNo
    Next colNo
    exportAnchor.Cells(1, 1).Resize(mTableShownCount + 1, mTableCols).Value = grid
    Set ExportTo = Me
End Function

' The rows the table shows under its filter, in order.
Public Property Get ShownRowCount() As Long
    RequireTable "ShownRowCount"
    EnsureTableShown
    ShownRowCount = mTableShownCount
End Property

Public Property Get RowCount() As Long
    RequireTable "RowCount"
    RowCount = mTableRows
End Property

' A cell's value by its row as the rows came in, whatever the sort.
Public Property Get CellValue(ByVal rowNumber As Long, ByVal columnNumber As Long) As Variant
    RequireTable "CellValue"
    If rowNumber < 1 Or rowNumber > mTableRows _
        Or columnNumber < 1 Or columnNumber > mTableCols Then
        RaiseError ERROR_INVALID_ARGUMENT, "CellValue", "No such cell."
    End If
    CellValue = mTableCells(columnNumber, rowNumber)
End Property

' The column the rows are sorted by, 0 for the order they came in.
Public Property Get SortColumn() As Long
    RequireTable "SortColumn"
    SortColumn = mSortColumn
End Property

Public Property Get SortedDescending() As Boolean
    RequireTable "SortedDescending"
    SortedDescending = mSortDescending
End Property

Private Sub RequireTable(ByVal memberName As String)
    RequireRole ROLE_COMPONENT, memberName
    If mKind <> KIND_TABLE Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            memberName & " applies to Table only."
    End If
End Sub

Private Sub RequireDatePicker(ByVal memberName As String)
    RequireRole ROLE_COMPONENT, memberName
    If mKind <> KIND_DATE Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            memberName & " applies to DatePicker only."
    End If
End Sub

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

' The alternative text screen readers announce for the control. ReDim
' writes one from the kind, text, and state ("Save, button", "Agree,
' checkbox, checked", ", unavailable" when disabled); AltText replaces
' it, and an empty string restores the generated one.
' An icon from the Windows icon font before the control's text, or alone
' when it has none: a name from ReDimUI.IconNames, such as "Save" or
' "Delete", or one character of IconFont; "" takes it off. An icon
' button without text takes the icon's name for screen readers.
Public Function Icon(ByVal iconName As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Icon"
    If mKind <> KIND_BUTTON And mKind <> KIND_LABEL Then
        RaiseError ERROR_INVALID_OPERATION, "Icon", _
            "Icon applies to Button and Label."
    End If
    If LenB(iconName) = 0 Then
        mIconGlyph = vbNullString
    Else
        mIconGlyph = ReDimUI.IconGlyph(iconName)
    End If
    mIconName = iconName
    Touch
    Set Icon = Me
End Function

' A count or short note in a pill on the control's top-right corner, in
' the danger color, such as the unread count on an Inbox button; ""
' takes it off. A click on it acts as a click on the control.
Public Function BadgeText(ByVal badgeWords As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BadgeText"
    mBadgeText = badgeWords
    Touch
    Set BadgeText = Me
End Function

' Puts the control in a stack, after the members already there. The
' stack places it from then on, so its own position no longer counts;
' its size still does. InStack "" takes it out where it stands.
Public Function InStack(ByVal stackId As String) As ReDimUI
    Dim hostApp As ReDimUI
    Dim stackValue As ReDimUI

    RequireRole ROLE_COMPONENT, "InStack"
    Set InStack = Me
    If StrComp(stackId, mStackId, vbTextCompare) = 0 Then Exit Function
    Set hostApp = OwnerApp
    If LenB(stackId) > 0 Then
        If Not hostApp.TryGetComponent(stackId, stackValue) Then
            RaiseError ERROR_NOT_FOUND, "InStack", _
                "No stack named '" & stackId & "' is in this app; create it first."
        End If
        If stackValue.ComponentKind <> KIND_STACK Then
            RaiseError ERROR_INVALID_ARGUMENT, "InStack", _
                "'" & stackId & "' is not a Stack."
        End If
    End If
    If LenB(mStackId) > 0 Then
        If hostApp.TryGetComponent(mStackId, stackValue) Then
            stackValue.LeaveStack mComponentId
        End If
    End If
    mStackId = stackId
    mAppliedFlowSig = vbNullString
    If LenB(stackId) > 0 Then
        If hostApp.TryGetComponent(stackId, stackValue) Then
            stackValue.JoinStack mComponentId
        End If
        mAnchor = vbNullString
        mRelMode = REL_NONE
    End If
    Touch
End Function

' Puts the control in an expander's panel: it shows only while the
' expander is open. The same as OnTab expanderId, 1.
Public Function InExpander(ByVal expanderId As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "InExpander"
    Set InExpander = OnTab(expanderId, 1)
End Function

' The points between a stack's members.
Public Function Gap(ByVal gapPoints As Double) As ReDimUI
    RequireStackKind "Gap"
    mStackGap = gapPoints
    MarkStackStale
    Touch
    Set Gap = Me
End Function

' The points between a stack's edges and its members.
Public Function Padding(ByVal paddingPoints As Double) As ReDimUI
    RequireStackKind "Padding"
    mStackPadding = paddingPoints
    MarkStackStale
    Touch
    Set Padding = Me
End Function

' Runs the members left to right instead of top to bottom, their tops
' aligned below the tallest caption among them.
Public Function Across(Optional ByVal acrossOn As Boolean = True) As ReDimUI
    RequireStackKind "Across"
    mStackAcross = acrossOn
    MarkStackStale
    Touch
    Set Across = Me
End Function

' A stack running down gives its members its own width, less padding.
Public Function Stretch(Optional ByVal stretchOn As Boolean = True) As ReDimUI
    RequireStackKind "Stretch"
    mStackStretch = stretchOn
    MarkStackStale
    Touch
    Set Stretch = Me
End Function

' Opens or closes an expander from code: its panel shows or hides, and
' nothing is written or fired.
Public Function Expanded(Optional ByVal expandedOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Expanded"
    If mKind <> KIND_EXPANDER Then
        RaiseError ERROR_INVALID_OPERATION, "Expanded", _
            "Expanded applies to Expander only."
    End If
    mChecked = expandedOn
    Touch
    Set Expanded = Me
End Function

Public Property Get IsExpanded() As Boolean
    RequireRole ROLE_COMPONENT, "IsExpanded"
    IsExpanded = (mKind = KIND_EXPANDER And mChecked)
End Property

' The palette entries a control offers: a shown, enabled button with a
' click handler offers its text; a menu offers each command as
' "Menu: Command".
Friend Sub OfferPaletteEntries(ByVal hostApp As ReDimUI, ByVal entryTexts As Collection)
    Dim commandWords As Variant

    If Not mVisible Or Not mEnabled Then Exit Sub
    If Left$(mComponentId, 4) = "mdl_" Then Exit Sub
    If mKind = KIND_BUTTON Then
        If LenB(mClickProc) > 0 Or LenB(mAsyncClickProc) > 0 Then
            hostApp.AddPaletteEntry entryTexts, mText, "btn" & vbTab & mComponentId
        End If
    ElseIf mIsMenu And Not mCommandOrder Is Nothing Then
        For Each commandWords In mCommandOrder
            If ItemPickable(ItemIndexOf(CStr(commandWords))) Then
                hostApp.AddPaletteEntry entryTexts, mText & ": " & CStr(commandWords), _
                    "menu" & vbTab & mComponentId & vbTab & CStr(commandWords)
            End If
        Next commandWords
    End If
End Sub

' The item number of an item's text, 0 when no item has it.
Private Function ItemIndexOf(ByVal itemText As String) As Long
    Dim idx As Long

    EnsureItemTexts
    For idx = 1 To ItemCount
        If StrComp(mItemTexts(idx), itemText, vbTextCompare) = 0 Then
            ItemIndexOf = idx
            Exit Function
        End If
    Next idx
End Function

Friend Sub MarkMenu()
    If mIsMenu Then Exit Sub
    mIsMenu = True
    If Not mVariantChosen Then mStyleVariant = VARIANT_SECONDARY
    Touch
End Sub

' A command. On a MenuButton it adds a row to the menu; on an app it adds
' an entry to the command palette. commandWords names it, handlerProc is
' the Public procedure it runs, and iconName an optional icon from
' IconNames. A menu command without a handler runs the menu's OnClick,
' where LastCommand names it.
Public Function AddCommand( _
    ByVal commandWords As String, _
    Optional ByVal handlerProc As String, _
    Optional ByVal iconName As String _
) As ReDimUI
    Dim commandKey As String
    Dim existing As Variant

    If mRole <> ROLE_APP Then
        RequireRole ROLE_COMPONENT, "AddCommand"
        If Not mIsMenu Then
            RaiseError ERROR_INVALID_OPERATION, "AddCommand", _
                "AddCommand applies to a MenuButton or an app."
        End If
    End If
    commandKey = LCase$(commandWords)
    If mCommandProcs Is Nothing Then
        Set mCommandProcs = New Collection
        Set mCommandIcons = New Collection
        Set mCommandOrder = New Collection
    End If
    If TryGetItem(mCommandProcs, commandKey, existing) Then
        mCommandProcs.Remove commandKey
        mCommandIcons.Remove commandKey
    Else
        mCommandOrder.Add commandWords
        If mRole <> ROLE_APP Then AddItem commandWords
    End If
    mCommandProcs.Add handlerProc, commandKey
    If LenB(iconName) > 0 Then
        mCommandIcons.Add ReDimUI.IconGlyph(iconName), commandKey
    Else
        mCommandIcons.Add vbNullString, commandKey
    End If
    Set AddCommand = Me
End Function

' The text of the command run last, on a MenuButton or an app.
Public Property Get LastCommand() As String
    LastCommand = mLastCommand
End Property

Private Function CommandProcOf(ByVal commandWords As String) As String
    Dim found As Variant

    If mCommandProcs Is Nothing Then Exit Function
    If TryGetItem(mCommandProcs, LCase$(commandWords), found) Then CommandProcOf = CStr(found)
End Function

Private Function CommandGlyphOf(ByVal commandWords As String) As String
    Dim found As Variant

    If mCommandIcons Is Nothing Then Exit Function
    If TryGetItem(mCommandIcons, LCase$(commandWords), found) Then CommandGlyphOf = CStr(found)
End Function

' A menu pick: the menu closes and the command's handler runs, with the
' menu as the sender; a command without one runs the menu's OnClick.
Private Sub RunMenuCommand(ByVal hostApp As ReDimUI, ByVal commandWords As String)
    Dim commandProc As String

    mSelectOpen = False
    mListHighlight = 0
    mLastCommand = commandWords
    Touch
    commandProc = CommandProcOf(commandWords)
    If LenB(commandProc) > 0 Then
        ReDimUI.RunHandlerAs hostApp, Me, commandProc
    Else
        RunClickHandlers hostApp
    End If
End Sub

' The palette runs a menu's command by its text.
Friend Sub RunCommandNamed(ByVal hostApp As ReDimUI, ByVal commandWords As String)
    RunMenuCommand hostApp, commandWords
End Sub

' A menu row's text: its icon in the icon font, then the command in the
' theme's, whatever font the replaced text began with.
Private Sub StyleMenuRow( _
    ByVal rowLetters As Object, _
    ByVal themeValue As ReDimUI, _
    ByVal rowText As String _
)
    rowLetters.Font.Name = themeValue.FontName
    If Left$(rowText, 1) <> vbTab And Mid$(rowText, 2, 1) = vbTab Then
        rowLetters.Characters(1, 1).Font.Name = ReDimUI.IconFont
    End If
End Sub

Private Sub RequireStackKind(ByVal memberName As String)
    RequireRole ROLE_COMPONENT, memberName
    If mKind <> KIND_STACK Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            memberName & " applies to Stack only."
    End If
End Sub

Public Function AltText(ByVal altDescription As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "AltText"
    mAltText = altDescription
    Touch
    Set AltText = Me
End Function

' A short note that shows once the pointer rests on the control for the
' Windows tooltip delay, and that screen readers read with the control.
Public Function Tooltip(ByVal tipText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "Tooltip"
    mTooltip = tipText
    If LenB(tipText) > 0 Then OwnerApp.NoteTooltips
    Touch
    Set Tooltip = Me
End Function

' The tooltip while the control is disabled: why it is, and what
' enables it.
Public Function DisabledReason(ByVal reasonText As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "DisabledReason"
    mDisabledReason = reasonText
    If LenB(reasonText) > 0 Then OwnerApp.NoteTooltips
    Touch
    Set DisabledReason = Me
End Function

Public Function Value(ByVal newNumber As Double) As ReDimUI
    RequireRole ROLE_COMPONENT, "Value"
    ' A group header is a label, never a selection.
    If mKind = KIND_SELECT Then
        If IsGroupHeader(CLng(newNumber)) Then newNumber = 0
    End If
    mValue = newNumber
    Touch
    Set Value = Me
End Function

Public Property Get CurrentValue() As Double
    RequireRole ROLE_COMPONENT, "CurrentValue"
    If mKind = KIND_TABS Then
        CurrentValue = ShownTab()
    Else
        CurrentValue = mValue
    End If
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
    ResetItemValues
    ForgetItemMarks
    ItemsChanged
    ClampItemSelection
    Touch
    Set Items = Me
End Function

' Replaces the item list from a 1D array, a Collection, a Range (one
' item per non-empty cell), or a ROneCOne sequence, so pickers can be
' fed straight from data. A value source of the same shape pairs each
' item with the value WritesTo writes for it.
Public Function ItemsFrom( _
    ByVal itemSource As Variant, _
    Optional ByVal valueSource As Variant _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "ItemsFrom"
    Set mItems = ReadItemSource(itemSource, "ItemsFrom")
    ResetItemValues
    ForgetItemMarks
    If Not IsMissing(valueSource) Then
        PairItemValues itemSource, valueSource, "ItemsFrom"
    End If
    ItemsChanged
    ClampItemSelection
    If Not mCheckedSet Is Nothing Then Set mCheckedSet = New Collection
    Touch
    Set ItemsFrom = Me
End Function

' Fills the transfer list's chosen (right-hand) side from a 1D array, a
' Collection, a Range, or a ROneCOne sequence, mirroring ItemsFrom for
' the available side, values included.
Public Function ChosenFrom( _
    ByVal itemSource As Variant, _
    Optional ByVal valueSource As Variant _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "ChosenFrom"
    Set mChosenItems = ReadItemSource(itemSource, "ChosenFrom")
    If Not IsMissing(valueSource) Then
        PairItemValues itemSource, valueSource, "ChosenFrom"
    End If
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

' The numbers a Sparkline plots, in order, from an array, a Collection,
' a Range (row by row), or a ROneCOne sequence; blanks and text are
' left out.
Public Function ValuesFrom(ByVal valueSource As Variant) As ReDimUI
    Dim gathered As Collection
    Dim entry As Variant
    Dim kept As Long

    RequireRole ROLE_COMPONENT, "ValuesFrom"
    If mKind <> KIND_SPARKLINE Then
        RaiseError ERROR_INVALID_OPERATION, "ValuesFrom", _
            "ValuesFrom applies to Sparkline only."
    End If
    Set gathered = ReadItemSource(valueSource, "ValuesFrom")
    Erase mSparkValues
    If gathered.Count > 0 Then
        ReDim mSparkValues(1 To gathered.Count)
        For Each entry In gathered
            If IsNumeric(entry) Then
                kept = kept + 1
                mSparkValues(kept) = CDbl(entry)
            End If
        Next entry
    End If
    mSparkCount = kept
    mSparkVersion = mSparkVersion + 1
    Touch
    Set ValuesFrom = Me
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

' Up and down arrows in the chosen panel's header move the chosen side's
' selected rows one place, as Alt+Up and Alt+Down do from its rows. The
' chosen order is part of the value, so a move writes WritesTo and
' fires OnChange.
Public Function Reorderable(Optional ByVal reorderOn As Boolean = True) As ReDimUI
    RequireRole ROLE_COMPONENT, "Reorderable"
    If mKind <> KIND_TRANSFER Then
        RaiseError ERROR_INVALID_OPERATION, "Reorderable", _
            "Reorderable applies to TransferList only."
    End If
    mReorderable = reorderOn
    Touch
    Set Reorderable = Me
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
' shifts their indices along with them. itemValue is what WritesTo
' writes when the item is picked; without one it writes the text.
Public Function AddItem( _
    ByVal itemText As String, _
    Optional ByVal atPosition As Long = 0, _
    Optional ByVal itemValue As Variant _
) As ReDimUI
    RequireRole ROLE_COMPONENT, "AddItem"
    If mItems Is Nothing Then Set mItems = New Collection
    If Not IsMissing(itemValue) Then StoreItemValue itemText, itemValue
    If atPosition >= 1 And atPosition <= mItems.Count Then
        mItems.Add itemText, Before:=atPosition
        If CLng(mValue) >= atPosition Then mValue = mValue + 1
        ShiftIndexesForInsert mCheckedSet, atPosition
        ShiftIndexesForInsert mDisabledItems, atPosition
        ShiftIndexesForInsert mGroupHeaders, atPosition
    Else
        mItems.Add itemText
    End If
    ItemsChanged
    Touch
    Set AddItem = Me
End Function

' A set of item numbers (checked rows, disabled items, group headers)
' follows its items through an insert or a removal.
Private Sub ShiftIndexesForInsert(ByVal indexSet As Collection, ByVal atPosition As Long)
    Dim idx As Long

    If indexSet Is Nothing Then Exit Sub
    For idx = 1 To indexSet.Count
        If CLng(indexSet.Item(idx)) >= atPosition Then
            indexSet.Add CLng(indexSet.Item(idx)) + 1, _
                Before:=idx
            indexSet.Remove idx + 1
        End If
    Next idx
End Sub

Private Sub ShiftIndexesForRemove(ByVal indexSet As Collection, ByVal removedAt As Long)
    Dim idx As Long

    If indexSet Is Nothing Then Exit Sub
    For idx = indexSet.Count To 1 Step -1
        If CLng(indexSet.Item(idx)) = removedAt Then
            indexSet.Remove idx
        ElseIf CLng(indexSet.Item(idx)) > removedAt Then
            indexSet.Add CLng(indexSet.Item(idx)) - 1, _
                Before:=idx
            indexSet.Remove idx + 1
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
    ' Read from the collection, not the snapshot: a removal leaves the
    ' snapshot stale, and rebuilding it for one text would walk the list.
    If Not mItemValueMap Is Nothing Then ForgetItemValue CStr(mItems.Item(idx))
    mItems.Remove idx
    ItemsChanged
    If CLng(mValue) = idx Then
        mValue = 0
    ElseIf CLng(mValue) > idx Then
        mValue = mValue - 1
    End If
    ShiftIndexesForRemove mCheckedSet, idx
    ShiftIndexesForRemove mDisabledItems, idx
    ShiftIndexesForRemove mGroupHeaders, idx
    Touch
    Set RemoveItem = Me
End Function

Public Function ClearItems() As ReDimUI
    RequireRole ROLE_COMPONENT, "ClearItems"
    Set mItems = New Collection
    Set mCommandProcs = Nothing
    Set mCommandIcons = Nothing
    Set mCommandOrder = Nothing
    ResetItemValues
    ForgetItemMarks
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

' A SelectBox item that shows but cannot be picked: it reads muted, the
' keys pass over it, and a click on it does nothing. A selection
' already on it stays.
Public Function ItemEnabled(ByVal itemNumber As Long, ByVal enabledFlag As Boolean) As ReDimUI
    RequireRole ROLE_COMPONENT, "ItemEnabled"
    RequireSelectKind "ItemEnabled"
    If itemNumber >= 1 And itemNumber <= ItemCount Then
        If mDisabledItems Is Nothing Then Set mDisabledItems = New Collection
        If TransferSelectionHas(mDisabledItems, itemNumber) = enabledFlag Then
            TransferSelectionToggle mDisabledItems, itemNumber
        End If
        Touch
    End If
    Set ItemEnabled = Me
End Function

Public Property Get IsItemEnabled(ByVal itemNumber As Long) As Boolean
    RequireRole ROLE_COMPONENT, "IsItemEnabled"
    IsItemEnabled = ItemPickable(itemNumber)
End Property

' Appends a group header to a SelectBox: a bold, muted row that labels
' the items after it. It takes a place in the item list, so ItemCount
' counts it and later item numbers count past it, and it cannot be
' picked.
Public Function AddGroup(ByVal groupCaption As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "AddGroup"
    RequireSelectKind "AddGroup"
    If mItems Is Nothing Then Set mItems = New Collection
    mItems.Add groupCaption
    If mGroupHeaders Is Nothing Then Set mGroupHeaders = New Collection
    mGroupHeaders.Add mItems.Count
    ItemsChanged
    Touch
    Set AddGroup = Me
End Function

Private Sub RequireSelectKind(ByVal memberName As String)
    If mKind <> KIND_SELECT Then
        RaiseError ERROR_INVALID_OPERATION, memberName, _
            memberName & " applies to SelectBox only."
    End If
End Sub

Private Sub ForgetItemMarks()
    Set mDisabledItems = Nothing
    Set mGroupHeaders = Nothing
End Sub

' True for an item that can be picked: in the list, and neither a
' disabled item nor a group header.
Private Function ItemPickable(ByVal itemNumber As Long) As Boolean
    If itemNumber < 1 Or itemNumber > ItemCount Then Exit Function
    If Not mDisabledItems Is Nothing Then
        If TransferSelectionHas(mDisabledItems, itemNumber) Then Exit Function
    End If
    If Not mGroupHeaders Is Nothing Then
        If TransferSelectionHas(mGroupHeaders, itemNumber) Then Exit Function
    End If
    ItemPickable = True
End Function

Private Function IsGroupHeader(ByVal itemNumber As Long) As Boolean
    If mGroupHeaders Is Nothing Then Exit Function
    IsGroupHeader = TransferSelectionHas(mGroupHeaders, itemNumber)
End Function

' The pickable item nearest a target in one direction, else in the
' other; 0 when nothing in the list can be picked.
Private Function PickableNear(ByVal aimItem As Long, ByVal stepDirection As Long) As Long
    Dim itemTotal As Long
    Dim probe As Long

    itemTotal = ItemCount
    If itemTotal = 0 Then Exit Function
    If aimItem < 1 Then aimItem = 1
    If aimItem > itemTotal Then aimItem = itemTotal
    If stepDirection = 0 Then stepDirection = 1
    For probe = aimItem To IIf(stepDirection > 0, itemTotal, 1) Step stepDirection
        If ItemPickable(probe) Then
            PickableNear = probe
            Exit Function
        End If
    Next probe
    For probe = aimItem To IIf(stepDirection > 0, 1, itemTotal) Step -stepDirection
        If ItemPickable(probe) Then
            PickableNear = probe
            Exit Function
        End If
    Next probe
End Function

Public Property Get ItemTextAt(ByVal itemNumber As Long) As String
    RequireRole ROLE_COMPONENT, "ItemTextAt"
    ItemTextAt = ItemAt(itemNumber)
End Property

' The item's value, or its text when it has none.
Public Property Get ItemValueAt(ByVal itemNumber As Long) As Variant
    Dim stateValue As Variant

    RequireRole ROLE_COMPONENT, "ItemValueAt"
    AssignVariant stateValue, ItemStateValue(ItemAt(itemNumber))
    If IsObject(stateValue) Then
        Set ItemValueAt = stateValue
    Else
        ItemValueAt = stateValue
    End If
End Property

' A transfer list's chosen item's value, or its text when it has none.
Public Property Get ChosenValueAt(ByVal itemNumber As Long) As Variant
    Dim stateValue As Variant

    RequireRole ROLE_COMPONENT, "ChosenValueAt"
    If mChosenItems Is Nothing Then Exit Property
    If itemNumber < 1 Or itemNumber > mChosenItems.Count Then Exit Property
    AssignVariant stateValue, ItemStateValue(CStr(mChosenItems.Item(itemNumber)))
    If IsObject(stateValue) Then
        Set ChosenValueAt = stateValue
    Else
        ChosenValueAt = stateValue
    End If
End Property

' What WritesTo writes for an item text: its value when it has one,
' else the text.
Private Function ItemStateValue(ByVal itemText As String) As Variant
    Dim found As Variant

    If Not mItemValueMap Is Nothing And LenB(itemText) > 0 Then
        If TryGetItem(mItemValueMap, itemText, found) Then
            If IsObject(found) Then
                Set ItemStateValue = found
            Else
                ItemStateValue = found
            End If
            Exit Function
        End If
    End If
    ItemStateValue = itemText
End Function

' An item's value as text, for the joined WritesTo of multi-pick lists;
' a value that has no text form writes the item's text.
Private Function ItemStateText(ByVal itemText As String) As String
    Dim stateValue As Variant

    If mItemValueMap Is Nothing Then
        ItemStateText = itemText
        Exit Function
    End If
    AssignVariant stateValue, ItemStateValue(itemText)
    ItemStateText = itemText
    If IsObject(stateValue) Or IsNull(stateValue) Then Exit Function
    On Error Resume Next
    ItemStateText = CStr(stateValue)
    On Error GoTo 0
End Function

Private Sub StoreItemValue(ByVal itemText As String, ByVal itemValue As Variant)
    If LenB(itemText) = 0 Then Exit Sub
    If mItemValueMap Is Nothing Then Set mItemValueMap = New Collection
    ForgetItemValue itemText
    mItemValueMap.Add itemValue, itemText
End Sub

Private Sub ForgetItemValue(ByVal itemText As String)
    If mItemValueMap Is Nothing Or LenB(itemText) = 0 Then Exit Sub
    On Error Resume Next
    mItemValueMap.Remove itemText
    On Error GoTo 0
End Sub

' A replaced item list drops its values, except those of a transfer
' list's chosen items, which the replacement leaves in place.
Private Sub ResetItemValues()
    Dim keptMap As Collection
    Dim entry As Variant
    Dim found As Variant

    If mItemValueMap Is Nothing Then Exit Sub
    If Not mChosenItems Is Nothing Then
        For Each entry In mChosenItems
            If TryGetItem(mItemValueMap, CStr(entry), found) Then
                If keptMap Is Nothing Then Set keptMap = New Collection
                On Error Resume Next
                keptMap.Add found, CStr(entry)
                On Error GoTo 0
            End If
        Next entry
    End If
    Set mItemValueMap = keptMap
End Sub

' Pairs item texts with values position by position, before empty
' entries are skipped, so a range of items with gaps still lines up
' with its range of values.
Private Sub PairItemValues( _
    ByVal itemSource As Variant, _
    ByVal valueSource As Variant, _
    ByVal memberName As String _
)
    Dim rawTexts As Collection
    Dim rawValues As Collection
    Dim pairedValues() As Variant
    Dim entry As Variant
    Dim idx As Long

    Set rawTexts = ReadRawSource(itemSource, memberName)
    Set rawValues = ReadRawSource(valueSource, memberName)
    If rawTexts.Count <> rawValues.Count Then
        RaiseError ERROR_INVALID_ARGUMENT, memberName, memberName & _
            " needs as many values as items."
    End If
    If rawTexts.Count = 0 Then Exit Sub
    ' Collection reads by index walk from the start; one pass each.
    ReDim pairedValues(1 To rawValues.Count)
    For Each entry In rawValues
        idx = idx + 1
        AssignVariant pairedValues(idx), entry
    Next entry
    idx = 0
    For Each entry In rawTexts
        idx = idx + 1
        If Not IsObject(entry) Then StoreItemValue CStr(entry), pairedValues(idx)
    Next entry
End Sub

' Every entry of a source, empty ones included, with its own type: the
' shape ReadItemSource reads, before it keeps only non-empty text.
Private Function ReadRawSource( _
    ByVal anySource As Variant, _
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
    If TypeName(anySource) = "Range" Then
        For Each area In anySource.Areas
            If area.CountLarge = 1 Then
                gathered.Add area.Value
            Else
                cellValues = area.Value
                For rowIdx = 1 To UBound(cellValues, 1)
                    For colIdx = 1 To UBound(cellValues, 2)
                        gathered.Add cellValues(rowIdx, colIdx)
                    Next colIdx
                Next rowIdx
            End If
        Next area
    ElseIf TypeName(anySource) = "Collection" Or TypeName(anySource) = "ROneCOne" Then
        For Each entry In anySource
            gathered.Add entry
        Next entry
    ElseIf IsArrayValue(anySource) Then
        For idx = LBound(anySource) To UBound(anySource)
            gathered.Add anySource(idx)
        Next idx
    Else
        RaiseError ERROR_INVALID_ARGUMENT, memberName, memberName & _
            " accepts an array, a Collection, a Range, or a ROneCOne sequence."
    End If
    Set ReadRawSource = gathered
End Function

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
    If LenB(mWritesToKey) > 0 Then
        If mKind = KIND_COMBO Then
            OwnerApp.SetState mWritesToKey, ItemStateValue(assignedText)
        Else
            OwnerApp.SetState mWritesToKey, assignedText
        End If
    End If
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
    If LenB(mWritesToKey) > 0 Then
        If mKind = KIND_COMBO Then
            hostApp.SetState mWritesToKey, ItemStateValue(cellText)
        Else
            hostApp.SetState mWritesToKey, cellText
        End If
    End If
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
    mToastActionBorder = -1
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
    mToastOriginTop = originTop
    mTop = originTop + hostApp.ToastStackAbove(mComponentId, slot)
    ' Entrance animation: the shape spawns below its slot and the pump
    ' eases it up into place. During a compaction glide the newcomer's
    ' slot may still be visually occupied by a survivor easing upward, so
    ' the entrance starts below the column as drawn, not just the model:
    ' a gap under the lowest live toast when that is deeper than the
    ' default hop.
    Dim occupiedBottom As Double
    mSlideRemaining = 14
    mToastExiting = False
    mExitRemaining = 0
    occupiedBottom = hostApp.MaxToastVisualBottom(mComponentId)
    If occupiedBottom >= 0 Then
        If occupiedBottom + TOAST_GAP - mTop > mSlideRemaining Then
            mSlideRemaining = occupiedBottom + TOAST_GAP - mTop
        End If
    End If
    If ReDimUI.MotionReduced Then mSlideRemaining = 0
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

' Moves this toast to a slot stackAbove points under the tray's top,
' easing up from wherever it currently sits (including mid-entrance)
' without a visual jump. A move down, when a toast above grew, lands at
' once, as does every move with motion reduced.
Friend Sub ReslotToast(ByVal newSlot As Long, ByVal stackAbove As Double)
    Dim visualTop As Double
    Dim targetTop As Double

    If mKind <> KIND_TOAST Then Exit Sub
    targetTop = mToastOriginTop + stackAbove
    If newSlot = CLng(mValue) And Abs(targetTop - mTop) < 0.01 Then Exit Sub
    visualTop = mTop + mSlideRemaining
    mValue = newSlot
    mTop = targetTop
    mSlideRemaining = visualTop - mTop
    If mSlideRemaining < 0 Then mSlideRemaining = 0
    If ReDimUI.MotionReduced Or mSlideRemaining = 0 Then
        ' No glide: the next apply moves the shape to its slot, and its
        ' buttons with it.
        mSlideRemaining = 0
        Touch
    End If
    ReDimUI.RequestPump
End Sub

Friend Property Get ToastHeight() As Double
    If mKind = KIND_TOAST Then ToastHeight = mHeight
End Property

' A toast under the pointer holds its countdown before the check.
Friend Function IsExpiredToast(ByVal nowMs As LongLong) As Boolean
    If mKind <> KIND_TOAST Then Exit Function
    HoldToastUnderPointer nowMs
    IsExpiredToast = (nowMs >= mExpiresAtMs)
End Function

' A toast under the pointer stops counting down, and keeps at least a
' second once the pointer leaves, so it can be read and acted on.
Private Sub HoldToastUnderPointer(ByVal nowMs As LongLong)
    Dim pointsX As Double
    Dim pointsY As Double

    If mKind <> KIND_TOAST Or mToastExiting Then Exit Sub
    If Not SheetInFront(OwnerApp) Then
        mToastHeldMs = 0
        Exit Sub
    End If
    If ReDimUI.FramePointer(pointsX, pointsY) Then
        If pointsX >= mLeft And pointsX <= mLeft + mWidth _
            And pointsY >= mTop And pointsY <= mTop + mHeight Then
            If mToastHeldMs = 0 Then
                mToastHeldMs = mExpiresAtMs - nowMs
                If mToastHeldMs < 1000 Then mToastHeldMs = 1000
            End If
            mExpiresAtMs = nowMs + mToastHeldMs
            Exit Sub
        End If
    End If
    mToastHeldMs = 0
End Sub

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
    ' Its panels show again, ungated.
    If mKind = KIND_TABS Or mKind = KIND_EXPANDER Then
        hostApp.ReapplyTabMembers mComponentId, -1
    End If
    ' The members after it in its stack close the gap.
    If LenB(mStackId) > 0 Then hostApp.ReflowStackNow mStackId
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

' True when a shape of this app's surely does not exist: the sheet held
' none of the app's shapes at mount and the app never drew this one.
' False only means the lookup has to be made.
Friend Function ShapeSurelyAbsent(ByVal drawnName As String) As Boolean
    Dim drawnMark As Variant

    If Not mSurfaceFresh Then Exit Function
    ShapeSurelyAbsent = Not TryGetItem(mDrawnShapes, drawnName, drawnMark)
End Function

Friend Sub NoteShapeDrawn(ByVal drawnName As String)
    Dim drawnMark As Variant

    If Not mSurfaceFresh Then Exit Sub
    If TryGetItem(mDrawnShapes, drawnName, drawnMark) Then Exit Sub
    mDrawnShapes.Add True, drawnName
End Sub

' A shape of this app's that was just deleted is surely absent again.
Friend Sub NoteShapeDropped(ByVal drawnName As String)
    Dim drawnMark As Variant

    If Not mSurfaceFresh Then Exit Sub
    If TryGetItem(mDrawnShapes, drawnName, drawnMark) Then mDrawnShapes.Remove drawnName
End Sub

' Whether any shape on the sheet has a name starting with namePrefix.
Private Function SheetHoldsPrefix(ByVal hostSheet As Worksheet, ByVal namePrefix As String) As Boolean
    Dim candidateShape As Shape
    Dim prefixLength As Long

    prefixLength = Len(namePrefix)
    For Each candidateShape In hostSheet.Shapes
        If Left$(candidateShape.Name, prefixLength) = namePrefix Then
            SheetHoldsPrefix = True
            Exit Function
        End If
    Next candidateShape
End Function

' A control started or stopped needing frames: the ticking list is built
' again the next time a frame asks for it.
Friend Sub TickablesChanged()
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
' take part, which keeps focus inside the dialog. With checkedFields the
' same order holds the fields ValidateAll checks instead.
Private Function TabOrder(Optional ByVal checkedFields As Boolean = False) As Collection
    Dim leading As Collection
    Dim trailing As Collection
    Dim orderedId As Variant
    Dim componentValue As ReDimUI
    Dim placed As ReDimUI
    Dim idx As Long
    Dim trapped As Boolean
    Dim inserted As Boolean
    Dim taking As Boolean

    Set leading = New Collection
    Set trailing = New Collection
    trapped = ModalUp()
    For Each orderedId In mComponentOrder
        If TryGetComponent(CStr(orderedId), componentValue) Then
            If checkedFields Then
                taking = componentValue.TakesChecks(Me)
            Else
                taking = componentValue.IsTabStop
                If taking And trapped Then
                    taking = (Left$(componentValue.ComponentId, 4) = "mdl_")
                End If
            End If
            If taking Then
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

' Checks the app's fields the way a commit does, Required first and
' then Validates: every enabled float field that takes a check and
' shows, or would once the Tabs and Expander panels it sits on show.
' Each that fails shows its message and rechecks as it is edited, and
' the first in Tab order gets focus, its panels shown first. True when
' every field passes.
Public Function ValidateAll() As Boolean
    Dim checkedFields As Collection
    Dim checkedField As ReDimUI
    Dim firstFailed As ReDimUI
    Dim idx As Long

    RequireRole ROLE_APP, "ValidateAll"
    ValidateAll = True
    Set checkedFields = TabOrder(True)
    If checkedFields.Count = 0 Then Exit Function
    BeginUpdate
    For idx = 1 To checkedFields.Count
        Set checkedField = checkedFields.Item(idx)
        If Not checkedField.PassesChecks(Me) Then
            ValidateAll = False
            If firstFailed Is Nothing Then Set firstFailed = checkedField
        End If
    Next idx
    If Not firstFailed Is Nothing Then
        firstFailed.RevealPanels Me
    End If
    EndUpdate
    If Not firstFailed Is Nothing Then
        firstFailed.Focus
    End If
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

' The shortcut codes this app's controls declare, in a Chr$(1)-marked list.
Friend Function ShortcutCodes() As String
    Dim componentValue As ReDimUI
    Dim keyCode As String

    For Each componentValue In mComponents
        keyCode = componentValue.ShortcutCode
        If LenB(keyCode) > 0 Then
            If Not CodeListHas(ShortcutCodes, keyCode) Then
                ShortcutCodes = CodeListAdd(ShortcutCodes, keyCode)
            End If
        End If
    Next componentValue
End Function

' Clicks the first control in creation order that declares the shortcut
' and can take a click; while a modal is up, only its buttons count.
Friend Function ActivateShortcut(ByVal keyCode As String) As Boolean
    Dim orderedId As Variant
    Dim componentValue As ReDimUI
    Dim trapped As Boolean

    RequireRole ROLE_APP, "ActivateShortcut"
    trapped = ModalUp()
    For Each orderedId In mComponentOrder
        If TryGetComponent(CStr(orderedId), componentValue) Then
            If StrComp(componentValue.ShortcutCode, keyCode, vbTextCompare) = 0 Then
                If componentValue.IsFocusable Then
                    If Not trapped Or Left$(componentValue.ComponentId, 4) = "mdl_" Then
                        ReDimUI.DispatchShape componentValue.ShapeName
                        ActivateShortcut = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next orderedId
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

    panelH = mHeight - TRANSFER_HEADER_H
    If panelH < 30 Then panelH = 30
    TransferRowCapacity = Int((panelH - 6) / TRANSFER_ROW_H)
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
' A list's item values in order, joined for a WritesTo state; texts for
' items without a value.
Private Function JoinedValues( _
    ByVal store As Collection, _
    ByVal joinWith As String _
) As String
    Dim parts() As String
    Dim entry As Variant
    Dim idx As Long

    If mItemValueMap Is Nothing Then
        JoinedValues = JoinedText(store, joinWith)
        Exit Function
    End If
    If store.Count = 0 Then Exit Function
    ReDim parts(1 To store.Count)
    For Each entry In store
        idx = idx + 1
        parts(idx) = ItemStateText(CStr(entry))
    Next entry
    JoinedValues = Join(parts, joinWith)
End Function

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

' Checked items' values (or texts) in list order, joined for the
' WritesTo state value.
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
            parts(takenCount) = ItemStateText(mItemTexts(idx))
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
    Dim movingFlags() As Boolean
    Dim shownItems() As Long
    Dim shownTotal As Long
    Dim keptSelection As Collection
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
    ReDim movingFlags(1 To sourceItems.Count)
    For Each entry In selectionSet
        If CLng(entry) >= 1 And CLng(entry) <= sourceItems.Count Then
            selectedFlags(CLng(entry)) = True
        End If
    Next entry
    ' A row the filter hides keeps its selection but stays put.
    shownTotal = TransferShownRows(IIf(rightward, 0, 1), shownItems)
    For idx = 1 To shownTotal
        If selectedFlags(shownItems(idx)) Then
            movingFlags(shownItems(idx)) = True
            anySelected = True
        End If
    Next idx
    idx = 0
    If Not anySelected Then Exit Sub
    ' One pass moves the selected rows to the target in their list order
    ' and rebuilds the source without them; removing by index would walk
    ' the list once per row. Selected rows the filter hid stay
    ' selected, renumbered.
    Set kept = New Collection
    Set keptSelection = New Collection
    For Each entry In sourceItems
        idx = idx + 1
        If movingFlags(idx) Then
            targetItems.Add CStr(entry)
        Else
            kept.Add CStr(entry)
            If selectedFlags(idx) Then keptSelection.Add kept.Count
        End If
    Next entry
    If rightward Then
        Set mItems = kept
        Set mSelLeft = keptSelection
    Else
        Set mChosenItems = kept
        Set mSelRight = keptSelection
    End If
    ItemsChanged
    Touch
    CommitTransferChange hostApp
End Sub

Private Sub TransferAll(ByVal hostApp As ReDimUI, ByVal rightward As Boolean)
    Dim entry As Variant
    Dim shownItems() As Long
    Dim shownTotal As Long
    Dim idx As Long

    ' A filtered side moves every row it shows, as a selection would.
    If LenB(PanelFilter(IIf(rightward, 0, 1))) > 0 Then
        shownTotal = TransferShownRows(IIf(rightward, 0, 1), shownItems)
        If shownTotal = 0 Then Exit Sub
        If rightward Then
            Set mSelLeft = New Collection
            For idx = 1 To shownTotal
                mSelLeft.Add shownItems(idx)
            Next idx
        Else
            Set mSelRight = New Collection
            For idx = 1 To shownTotal
                mSelRight.Add shownItems(idx)
            Next idx
        End If
        TransferSelected hostApp, rightward
        Exit Sub
    End If
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
        hostApp.SetState mWritesToKey, JoinedValues(mChosenItems, ", ")
    End If
    If LenB(mClickProc) > 0 Then
        On Error Resume Next
        Application.Run mClickProc
        On Error GoTo 0
    End If
End Sub

Friend Function CanClick() As Boolean
    If Not mVisible Or Not mEnabled Or mBusy Then Exit Function
    ' A float field and a transfer list take every click, so the second
    ' of a double click reaches them: it selects a field's word or moves
    ' a transfer row across. A table's second click on a header flips the
    ' sort, and an open calendar takes quick clicks on its month arrows.
    If Not IsFloatField() And mKind <> KIND_TRANSFER And mKind <> KIND_TABLE _
        And Not (mKind = KIND_DATE And mSelectOpen) Then
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
    Dim clickedPart As String

    ' The click Excel delivers for the release of a held press that
    ' already repeated.
    If GetTickCount64() < mHoldSuppressUntilMs Then Exit Sub
    Set hostApp = OwnerApp
    clickedPart = ReDimUI.SenderPart
    Select Case mKind
        Case KIND_TOGGLE
            mChecked = Not mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
            Touch
        Case KIND_INPUT
            If clickedPart = "cx" Then
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
            If clickedPart = "cx" Then
                ClearFieldText hostApp
                Exit Sub
            ElseIf clickedPart = "optn" Then
                ' The empty row is a message, not an option.
                Exit Sub
            ElseIf clickedPart = "optu" Or clickedPart = "optd" Then
                PageDropList (clickedPart = "optd")
                Exit Sub
            ElseIf Left$(clickedPart, 3) = "opt" _
                And IsNumeric(Mid$(clickedPart, 4)) Then
                ' Clicked window rows map through the scroll offset to
                ' their absolute match index. A pick that leaves the value
                ' as it was runs no handlers.
                If Not PickComboItem(hostApp, _
                    mListOffset + CLng(Mid$(clickedPart, 4))) Then
                    Exit Sub
                End If
            ElseIf IsFloatField() Then
                ' The caret glyph toggles the list; the face takes the
                ' caret like any field.
                If clickedPart = "caret" _
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
            If clickedPart = "optn" Then
                Exit Sub
            ElseIf clickedPart = "optu" Or clickedPart = "optd" Then
                PageDropList (clickedPart = "optd")
                Exit Sub
            ElseIf Left$(clickedPart, 3) = "opt" _
                And IsNumeric(Mid$(clickedPart, 4)) Then
                Dim pickedItem As Long
                pickedItem = ListItemAtRow(CLng(Mid$(clickedPart, 4)))
                If Not ItemPickable(pickedItem) Then Exit Sub
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
            checkPart = clickedPart
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
            transferPart = clickedPart
            EnsureTransferCollections
            If transferPart = "alu" Or transferPart = "ald" _
                Or transferPart = "clu" Or transferPart = "cld" Then
                PageTransferPanel transferPart
            ElseIf Left$(transferPart, 2) = "al" Then
                If IsNumeric(Mid$(transferPart, 3)) Then
                    ' Window rows map through the scroll offset and the
                    ' filter.
                    transferRow = TransferShownItem(0, mLeftOffset + CLng(Mid$(transferPart, 3)))
                    If transferRow = 0 Then Exit Sub
                    If RowDoubleClicked(transferPart) Then
                        MoveTransferRow hostApp, True, transferRow
                    Else
                        TransferSelectionToggle mSelLeft, transferRow
                        PlaceKeyCursor 0, transferRow
                        Touch
                    End If
                End If
            ElseIf Left$(transferPart, 2) = "cl" Then
                If IsNumeric(Mid$(transferPart, 3)) Then
                    transferRow = TransferShownItem(1, mRightOffset + CLng(Mid$(transferPart, 3)))
                    If transferRow = 0 Then Exit Sub
                    If RowDoubleClicked(transferPart) Then
                        MoveTransferRow hostApp, False, transferRow
                    Else
                        TransferSelectionToggle mSelRight, transferRow
                        PlaceKeyCursor 1, transferRow
                        Touch
                    End If
                End If
            ElseIf transferPart = "mvr" Then
                TransferSelected hostApp, True
            ElseIf transferPart = "mvl" Then
                TransferSelected hostApp, False
            ElseIf transferPart = "mvar" Then
                TransferAll hostApp, True
            ElseIf transferPart = "mval" Then
                TransferAll hostApp, False
            ElseIf transferPart = "mvu" Or transferPart = "mvd" Then
                ReorderChosen hostApp, (transferPart = "mvu")
            End If
            Exit Sub
        Case KIND_DATE
            If Not ClickCalendar(hostApp, clickedPart) Then Exit Sub
        Case KIND_TABLE
            If Not ClickTable(hostApp, clickedPart) Then Exit Sub
        Case KIND_EXPANDER
            ToggleExpander hostApp
        Case KIND_TABS
            Dim tabPicked As Long
            If Left$(clickedPart, 2) = "tb" _
                And IsNumeric(Mid$(clickedPart, 3)) Then
                tabPicked = CLng(Mid$(clickedPart, 3))
            End If
            If Not SelectTab(hostApp, tabPicked) Then Exit Sub
        Case KIND_RADIO
            Dim rowNumber As Long
            If LenB(clickedPart) = 0 Then
                rowNumber = 1
            ElseIf IsNumeric(Mid$(clickedPart, 2)) Then
                rowNumber = CLng(Mid$(clickedPart, 2))
            End If
            If Not SelectRadioRow(hostApp, rowNumber) Then Exit Sub
        Case KIND_STEPPER
            Dim stepped As Double
            If clickedPart = "minus" Then
                stepped = mValue - mSliderStep
            ElseIf clickedPart = "plus" Then
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
            ' Any click puts the toast away at once, fading from here: a
            ' dismissal left to the countdown was undone while the pointer
            ' rested on the toast, which holds the countdown. The action
            ' button runs its handler and the card runs OnClick; the close
            ' button runs nothing. Handler errors reach the app's handler.
            BeginToastExit
            Select Case clickedPart
                Case "ta"
                    If LenB(mToastActionProc) > 0 Then Application.Run mToastActionProc
                Case "tx"
                Case Else
                    RunClickHandlers hostApp
            End Select
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
    mItemsVersion = mItemsVersion + 1
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
    ' A member's change may move the members after it: its stack lays
    ' out again the next time any of them asks for a slot.
    If LenB(mStackId) > 0 Then hostApp.NoteStackTouched mStackId
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
    ' Feature steps run only for controls that use the feature: VBA pays
    ' for each procedure's first call in a session, and a cold build
    ' paid for steps that had nothing to do.
    If LenB(mTabOwnerId) > 0 Or LenB(mStackId) > 0 Or mTabShut Then
        RefreshTabShut hostApp
    End If
    If mKind = KIND_EXPANDER Then
        If mChecked Then
            mIconGlyph = ReDimUI.IconGlyph("ChevronDown")
        Else
            mIconGlyph = ReDimUI.IconGlyph("ChevronRight")
        End If
    End If
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
    If LenB(mCaptionText) > 0 Or LenB(mAppliedCaptionSig) > 0 _
        Or LenB(mHintText) > 0 Or LenB(mAppliedHintSig) > 0 Then
        ApplyAdornments hostSheet, themeValue
    End If
    If LenB(mBadgeText) > 0 Or LenB(mAppliedBadgeSig) > 0 Then
        ApplyBadgePart hostSheet, themeValue
    End If
    If LenB(mAppliedRingSig) > 0 _
        Or ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
        ApplyFocusRing hostSheet, themeValue
    End If
    ApplyAltText mainShape
    mAppliedValid = True
    ' A field that grew or shrank since it was drawn moves what is placed
    ' against it; a first draw leaves that to the render in progress.
    If wasApplied And Abs(mGrowAdded - growBefore) > 0.01 Then
        hostApp.ReapplyDependents mComponentId
    End If
    If mKind = KIND_TABS Or mKind = KIND_EXPANDER Then SyncTabPanels hostApp
    If mKind = KIND_STACK Then MoveStackMembers hostSheet, themeValue, hostApp
    If LenB(mStackId) > 0 Then NoteFlow hostApp, themeValue
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
            If LenB(mText) = 0 And Len(mIconName) > 1 Then
                described = Phrase(UI_ALT_BUTTON, mIconName)
            Else
                described = Phrase(UI_ALT_BUTTON, mText)
            End If
            If mBusy Then described = described & ", " & Phrase(UI_ALT_BUSY)
        Case KIND_PROGRESS
            described = Phrase(UI_ALT_PROGRESS, CLng(mValue / 5) * 5)
        Case KIND_SPINNER
            described = Phrase(UI_ALT_SPINNER)
        Case KIND_TOGGLE
            If mChecked Then
                described = Phrase(UI_ALT_SWITCH_ON)
            Else
                described = Phrase(UI_ALT_SWITCH_OFF)
            End If
        Case KIND_INPUT
            described = Phrase(UI_ALT_EDIT_FIELD)
            If mMasked Then described = described & ", " & Phrase(UI_ALT_MASKED)
            If LenB(mPlaceholder) > 0 Then described = described & ", " & mPlaceholder
        Case KIND_COMBO
            described = Phrase(UI_ALT_COMBO_BOX)
            If LenB(mPlaceholder) > 0 Then described = described & ", " & mPlaceholder
        Case KIND_SELECT
            If mIsMenu Then
                described = Phrase(UI_ALT_MENU_BUTTON, mText)
            ElseIf CLng(mValue) >= 1 And CLng(mValue) <= ItemCount Then
                described = Phrase(UI_ALT_DROP_DOWN, ItemAt(CLng(mValue)))
            Else
                described = Phrase(UI_ALT_DROP_DOWN, mText)
            End If
        Case KIND_TICKBOX
            If mChecked Then
                described = Phrase(UI_ALT_CHECKED, mText)
            Else
                described = Phrase(UI_ALT_NOT_CHECKED, mText)
            End If
        Case KIND_RADIO
            described = Phrase(UI_ALT_RADIO_GROUP)
            If CLng(mValue) >= 1 And CLng(mValue) <= ItemCount Then
                described = described & ", " & Phrase(UI_ALT_ITEM_SELECTED, ItemAt(CLng(mValue)))
            End If
        Case KIND_STEPPER
            described = Phrase(UI_ALT_STEPPER, mValue)
        Case KIND_SLIDEBAR
            described = Phrase(UI_ALT_SLIDER, mValue)
        Case KIND_TRANSFER
            described = Phrase(UI_ALT_TRANSFER_LIST)
        Case KIND_CHECKLIST
            EnsureCheckedSet
            described = Phrase(UI_ALT_CHECK_LIST, mCheckedSet.Count, ItemCount)
        Case KIND_IMAGE
            described = Phrase(UI_ALT_IMAGE)
        Case KIND_SKELETON
            described = Phrase(UI_ALT_LOADING)
        Case KIND_SPARKLINE
            described = SparklineSummary()
        Case KIND_EXPANDER
            If mChecked Then
                described = Phrase(UI_ALT_EXPANDED, mText)
            Else
                described = Phrase(UI_ALT_COLLAPSED, mText)
            End If
        Case KIND_TABLE
            If mTableRows = 1 Then
                described = Phrase(UI_ALT_TABLE_ONE_ROW)
            Else
                described = Phrase(UI_ALT_TABLE, mTableRows)
            End If
            If mSortColumn > 0 Then
                If mSortDescending Then
                    described = described & ", " & _
                        Phrase(UI_ALT_SORTED_DESCENDING, mTableHeads(mSortColumn))
                Else
                    described = described & ", " & _
                        Phrase(UI_ALT_SORTED_ASCENDING, mTableHeads(mSortColumn))
                End If
            End If
            If LenB(mTableFilter) > 0 Then
                EnsureTableShown
                described = described & ", " & Phrase(UI_ALT_SHOWN, mTableShownCount)
            End If
            If CLng(mValue) >= 1 And CLng(mValue) <= mTableRows Then
                described = described & ", " & Phrase(UI_ALT_ITEM_SELECTED, _
                    TableCellText(1, mTableCells(1, CLng(mValue))))
            End If
        Case KIND_DATE
            If mValue <> 0 Then
                described = Phrase(UI_ALT_DATE_PICKER, FormatDateTime(CDate(mValue), vbLongDate))
            Else
                described = Phrase(UI_ALT_DATE_PICKER, mText)
            End If
        Case KIND_TABS
            described = Phrase(UI_ALT_TABS)
            If ShownTab() >= 1 Then
                described = described & ", " & _
                    Phrase(UI_ALT_TAB_SELECTED, ItemAt(ShownTab()), ShownTab(), ItemCount)
            End If
    End Select
    If Not mEnabled And LenB(described) > 0 Then
        described = described & ", " & Phrase(UI_ALT_UNAVAILABLE)
    End If
    If LenB(mAccessKey) > 0 And LenB(described) > 0 Then
        described = described & ", Alt+" & UCase$(mAccessKey)
    End If
    If LenB(mShortcut) > 0 And LenB(described) > 0 Then
        described = described & ", " & ShortcutLabel(mShortcut)
    End If
    ' The tooltip is read with the control: the reason while disabled.
    If LenB(described) > 0 Then
        If Not mEnabled And LenB(mDisabledReason) > 0 Then
            described = described & ", " & mDisabledReason
        ElseIf LenB(mTooltip) > 0 Then
            described = described & ", " & mTooltip
        End If
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
            If mKind = KIND_DATE Then
                ' A Date is not numeric to VBA; a date picker takes either.
                If IsDate(bound) Or IsNumeric(bound) Then
                    mValue = CDbl(DayOnly(CDate(bound)))
                End If
            ElseIf mKind = KIND_TOGGLE Or mKind = KIND_TICKBOX Or mKind = KIND_EXPANDER Then
                ' A switch, a tick box, or an expander follows True or False.
                If VarType(bound) = vbBoolean Or IsNumeric(bound) Then
                    mChecked = CBool(bound)
                End If
            ElseIf IsNumeric(bound) Then
                mValue = CDbl(bound)
            End If
        End If
    End If
    If LenB(mBindVisibleKey) > 0 Then
        If hostApp.HasState(mBindVisibleKey) Then
            If mTabShut Then
                mVisibleBeforeTab = CBool(hostApp.State(mBindVisibleKey)) _
                    Xor mBindVisibleInvert
            Else
                mVisible = CBool(hostApp.State(mBindVisibleKey)) Xor mBindVisibleInvert
            End If
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

' A control's rectangle: its stack's slot for it, or its own placement;
' then a stack lays its members out, which sizes it.
Private Sub ResolveGeometry(ByVal hostSheet As Worksheet)
    Dim stacked As Boolean

    If LenB(mStackId) > 0 Then stacked = PlaceInStack(hostSheet)
    If Not stacked Then ResolvePlacement hostSheet
    If mKind = KIND_STACK Then ArrangeStack
End Sub

' True when the control's stack placed it: at its slot, or where it stood
' while hidden. A stack that no longer exists places nothing.
Private Function PlaceInStack(ByVal hostSheet As Worksheet) As Boolean
    Dim stackValue As ReDimUI
    Dim slotLeft As Double
    Dim slotTop As Double
    Dim slotSpan As Double

    If LenB(mStackId) = 0 Then Exit Function
    If Not OwnerApp.TryGetComponent(mStackId, stackValue) Then Exit Function
    stackValue.EnsureGeometryResolved hostSheet
    If stackValue.StackSlot(mComponentId, slotLeft, slotTop, slotSpan) Then
        mLeft = slotLeft
        mTop = slotTop
        If slotSpan > 0 Then mWidth = slotSpan
    End If
    PlaceInStack = True
End Function

Private Sub ResolvePlacement(ByVal hostSheet As Worksheet)
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
    Dim hostApp As ReDimUI

    Set hostApp = OwnerApp
    If Not hostApp.ShapeSurelyAbsent(ShapeName) Then
        On Error Resume Next
        Set mainShape = hostSheet.Shapes(ShapeName)
        On Error GoTo 0
    End If
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
        hostApp.NoteShapeDrawn ShapeName
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
        Case KIND_SELECT, KIND_DATE
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
        Case KIND_SKELETON
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
            mainShape.Adjustments(1) = 0.1
        Case KIND_BADGE
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, BadgeSpan(), mHeight)
            mainShape.Adjustments(1) = 0.5
        Case KIND_STACK
            ' Behind the controls it holds, whenever it was made.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRectangle, mLeft, mTop, IIf(mWidth > 0, mWidth, 1), _
                IIf(mHeight > 0, mHeight, 1))
            mainShape.ZOrder msoSendToBack
        Case KIND_TABS
            ' The rule along the strip's foot; the tabs are parts.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRectangle, mLeft, mTop + mHeight - 1, mWidth, 1)
        Case KIND_EXPANDER
            ' A header whose hover fill rounds its corners.
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRoundedRectangle, mLeft, mTop, mWidth, mHeight)
        Case Else
            Set mainShape = hostSheet.Shapes.AddShape( _
                msoShapeRectangle, mLeft, mTop, mWidth, mHeight)
    End Select
    mainShape.Placement = xlFreeFloating
    Set CreateShapeForKind = mainShape
End Function

' Gives a part the rounded corners of a list row or hover fill, a radius
' of ROW_CORNER points on a part partHeight tall. A plain rectangle left
' by a build before rounded rows turns rounded here.
Private Sub RoundCorners(ByVal partShape As Shape, ByVal partHeight As Double)
    If partShape.AutoShapeType <> msoShapeRoundedRectangle Then
        partShape.AutoShapeType = msoShapeRoundedRectangle
    End If
    If partHeight > 0 Then partShape.Adjustments(1) = ROW_CORNER / partHeight
End Sub

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
    Dim hostApp As ReDimUI

    partShapeName = ShapeName & PART_SEPARATOR & partName
    Set hostApp = OwnerApp
    If Not hostApp.ShapeSurelyAbsent(partShapeName) Then
        On Error Resume Next
        Set partShape = hostSheet.Shapes(partShapeName)
        On Error GoTo 0
    End If
    wasCreated = (partShape Is Nothing)
    If wasCreated Then
        Set partShape = hostSheet.Shapes.AddShape(shapeType, mLeft, mTop, 8, 8)
        partShape.Name = partShapeName
        partShape.Placement = xlFreeFloating
        partShape.Line.Visible = msoFalse
        hostApp.NoteShapeDrawn partShapeName
    End If
    Set EnsurePartShape = partShape
End Function

' A new part made as a copy of a part already drawn in the look it needs:
' the copy carries the size, fill, border, margins, tab stops, font, and
' click, for a fraction of writing them. Nothing when a shape by that name
' exists, so the caller draws that one in full, or when the copy fails. A
' sheet protected by its host refuses copies, and its parts draw in full.
Private Function CopyPartShape( _
    ByVal hostSheet As Worksheet, _
    ByVal templateShape As Shape, _
    ByVal partName As String _
) As Shape
    Dim partShape As Shape
    Dim partShapeName As String
    Dim hostApp As ReDimUI
    Dim protectionLifted As Boolean

    partShapeName = ShapeName & PART_SEPARATOR & partName
    Set hostApp = OwnerApp
    If Not hostApp.ShapeSurelyAbsent(partShapeName) Then
        On Error Resume Next
        Set partShape = hostSheet.Shapes(partShapeName)
        On Error GoTo 0
        If Not partShape Is Nothing Then Exit Function
    End If
    If hostSheet.ProtectDrawingObjects Then
        protectionLifted = hostApp.LiftSurfaceProtection()
        If Not protectionLifted Then Exit Function
    End If
    On Error Resume Next
    ' Excel's type library declares Shape.Duplicate As Shape; the analyzer's
    ' Excel model says ShapeRange, from XLIDE's data through 10.7.0.
    '@pyvba-ignore-next-line: assignment-object-type-mismatch
    Set partShape = templateShape.Duplicate
    On Error GoTo 0
    If protectionLifted Then hostApp.RestoreSurfaceProtection
    If partShape Is Nothing Then Exit Function
    partShape.Name = partShapeName
    hostApp.NoteShapeDrawn partShapeName
    Set CopyPartShape = partShape
End Function

' Deletes a part that may exist. One this app never drew on a sheet that
' held none of its shapes cannot, and looking for it would cost a miss.
' A part deleted here is surely absent again, so a list or calendar that
' opens anew draws its parts without a lookup that could only miss.
Private Sub DropPart(ByVal hostSheet As Worksheet, ByVal partName As String)
    Dim partShapeName As String
    Dim hostApp As ReDimUI

    partShapeName = ShapeName & PART_SEPARATOR & partName
    Set hostApp = OwnerApp
    If hostApp.ShapeSurelyAbsent(partShapeName) Then Exit Sub
    On Error Resume Next
    hostSheet.Shapes(partShapeName).Delete
    If Err.Number = 0 Then hostApp.NoteShapeDropped partShapeName
    On Error GoTo 0
End Sub

Private Function IsCompositeDrawnKind() As Boolean
    IsCompositeDrawnKind = (mKind = KIND_TICKBOX Or mKind = KIND_RADIO _
        Or mKind = KIND_STEPPER Or mKind = KIND_SLIDEBAR _
        Or mKind = KIND_TRANSFER Or mKind = KIND_CHECKLIST _
        Or mKind = KIND_TABS)
End Function

' Geometry diffs against the shape's actual position rather than a cache, so
' a shape moved by hand (Ctrl+drag) snaps back to its declared rectangle on
' the next render. Composite drawn kinds own their geometry in extras: the
' main shape is one piece of several, not the declared rectangle.
Private Sub ApplyGeometry(ByVal mainShape As Shape)
    If IsCompositeDrawnKind() Then Exit Sub
    If Abs(mainShape.Left - mLeft) > 0.01 Then mainShape.Left = mLeft
    If Abs(mainShape.Top - mTop) > 0.01 Then mainShape.Top = mTop
    If mKind = KIND_BADGE Then
        If Abs(mainShape.Width - BadgeSpan()) > 0.01 Then mainShape.Width = BadgeSpan()
    ElseIf Abs(mainShape.Width - mWidth) > 0.01 Then
        mainShape.Width = mWidth
    End If
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
        Case KIND_LABEL, KIND_SPARKLINE
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
        Case KIND_STACK
            ' A stack is clear unless given a fill, and has no edge.
            If Not mAppliedValid Or mAppliedFill <> mFillColor Then
                mainShape.Fill.Visible = BoolToMso(mFillColor >= 0)
                If mFillColor >= 0 Then mainShape.Fill.ForeColor.RGB = mFillColor
                mainShape.Line.Visible = msoFalse
                mAppliedFill = mFillColor
            End If
        Case KIND_EXPANDER
            ' A header row, clear until the pointer tints it.
            resolvedFill = PointerTint(themeValue.SurfaceColor, themeValue.OnSurfaceColor)
            If resolvedFill = themeValue.SurfaceColor Then resolvedFill = -1
            If Not mAppliedValid Then RoundCorners mainShape, mHeight
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = BoolToMso(resolvedFill >= 0)
                If resolvedFill >= 0 Then mainShape.Fill.ForeColor.RGB = resolvedFill
                mainShape.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
        Case KIND_BADGE
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = resolvedFill
                mainShape.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
        Case KIND_CARD, KIND_TOAST, KIND_TABLE
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
            resolvedFill = PointerTint(ResolveToggleFill(themeValue), _
                IIf(mChecked, themeValue.OnPrimaryColor, themeValue.OnSurfaceColor))
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = resolvedFill
                mainShape.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
        Case KIND_IMAGE
            ApplyImageFill mainShape, themeValue
        Case KIND_SKELETON
            ' With lines the bars carry the pulse and the block is clear.
            If mSkeletonLines > 0 Then
                resolvedFill = -1
            Else
                resolvedFill = SkeletonFill(themeValue)
            End If
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = BoolToMso(resolvedFill >= 0)
                If resolvedFill >= 0 Then mainShape.Fill.ForeColor.RGB = resolvedFill
                mainShape.Line.Visible = msoFalse
                mAppliedFill = resolvedFill
            End If
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
                errorOn = LenB(ShownError()) > 0
                resolvedFill = PointerTint(ResolveSurfaceFill(themeValue), _
                    themeValue.OnSurfaceColor)
                If Not mAppliedValid _
                    Or mAppliedFill <> resolvedFill _
                    Or mAppliedFocusRing <> ringOn _
                    Or mAppliedErrorBorder <> errorOn Then
                    mainShape.Fill.Visible = msoTrue
                    mainShape.Fill.ForeColor.RGB = resolvedFill
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
                    mAppliedFill = resolvedFill
                    mAppliedFocusRing = ringOn
                End If
            ElseIf Not mAppliedValid Or mAppliedFill <> themeValue.BorderColor Then
                mainShape.Fill.Visible = msoFalse
                mainShape.Line.Visible = msoTrue
                mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                mainShape.Line.Weight = 1
                mAppliedFill = themeValue.BorderColor
            End If
        Case KIND_SELECT, KIND_DATE
            ' A menu's face is a button in its variant's colors.
            If mIsMenu Then
                resolvedFill = PointerTint(ResolveFill(themeValue), ResolveTextColor(themeValue))
            Else
                resolvedFill = PointerTint(ResolveSurfaceFill(themeValue), _
                    themeValue.OnSurfaceColor)
            End If
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = resolvedFill
                mainShape.Line.Visible = msoTrue
                mainShape.Line.ForeColor.RGB = themeValue.BorderColor
                mainShape.Line.Weight = 1
                mAppliedFill = resolvedFill
            End If
        Case KIND_OVERLAY
            If Not mAppliedValid Then
                mainShape.Fill.Visible = msoTrue
                mainShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
                mainShape.Fill.Transparency = 0.45
                mainShape.Line.Visible = msoFalse
            End If
        Case Else
            resolvedFill = PointerTint(resolvedFill, ResolveTextColor(themeValue))
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
    StyleGlyphText mainShape, Phrase(UI_IMAGE_NOT_FOUND), themeValue.OnMutedColor, _
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
    ' A menu's face inks like a button's.
    If mIsMenu Then
        If mStyleVariant = VARIANT_SECONDARY Or mStyleVariant = VARIANT_NEUTRAL Then
            ResolveTextColor = themeValue.OnSurfaceColor
        Else
            ResolveTextColor = themeValue.OnPrimaryColor
        End If
        Exit Function
    End If
    ' A date picker's placeholder reads muted, as a field's does.
    If mKind = KIND_DATE And mValue = 0 Then
        ResolveTextColor = themeValue.OnMutedColor
        Exit Function
    End If
    Select Case mKind
        Case KIND_LABEL, KIND_CARD, KIND_SELECT, KIND_TOAST, _
            KIND_INPUT, KIND_COMBO, KIND_DATE, KIND_EXPANDER
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
    Dim busyWords As String

    If mBusy And mKind = KIND_BUTTON Then
        If mBusyTextGiven Then
            busyWords = mBusyText
        Else
            busyWords = Phrase(UI_WORKING)
        End If
        If LenB(busyWords) > 0 Then
            ResolveText = busyWords
            Exit Function
        End If
    End If
    If mKind = KIND_SELECT Then
        ResolveText = ItemAt(CLng(mValue))
        If LenB(ResolveText) = 0 Then ResolveText = mText
        Exit Function
    End If
    If mKind = KIND_DATE Then
        If mValue = 0 Then
            ResolveText = mText
        ElseIf LenB(mDateFormat) > 0 Then
            ResolveText = Format$(CDate(mValue), mDateFormat)
        Else
            ResolveText = FormatDateTime(CDate(mValue), vbShortDate)
        End If
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
    ' A masked field renders a dot for each character, one for one, so
    ' the caret and the face map line up with the text beneath.
    If mMasked Then
        rendered = String$(mCaretPos, ChrW(8226)) & IIf(mCaretOn, "|", " ") & _
            String$(Len(mText) - mCaretPos, ChrW(8226)) & ghostText
    Else
        rendered = Left$(mText, mCaretPos) & IIf(mCaretOn, "|", " ") & _
            Mid$(mText, mCaretPos + 1) & ghostText
    End If
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
        Or mKind = KIND_IMAGE Or mKind = KIND_SKELETON Or mKind = KIND_TABLE _
        Or mKind = KIND_STACK Or mKind = KIND_SPARKLINE _
        Or IsCompositeDrawnKind() Then Exit Sub
    If (mKind = KIND_INPUT Or mKind = KIND_COMBO) And Not IsFloatField() Then
        Exit Sub
    End If
    shownText = ResolveText(Nothing)
    ' An icon leads the text, or stands alone.
    If LenB(mIconGlyph) > 0 Then
        If LenB(shownText) > 0 Then
            shownText = mIconGlyph & "  " & shownText
        Else
            shownText = mIconGlyph
        End If
    End If
    resolvedColor = ResolveTextColor(themeValue)
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then
        resolvedSize = themeValue.BaseFontSize
        If mKind = KIND_BADGE Then resolvedSize = resolvedSize - 2
    End If
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
        ElseIf mMasked Then
            shownText = String$(Len(mText), ChrW(8226))
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
        If mKind = KIND_BADGE Then
            mainShape.TextFrame2.MarginLeft = 4
            mainShape.TextFrame2.MarginRight = 4
            mainShape.TextFrame2.MarginTop = 0
            mainShape.TextFrame2.MarginBottom = 0
            mainShape.TextFrame2.WordWrap = msoFalse
        End If
        If mKind = KIND_BUTTON Or mKind = KIND_BADGE Then
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
        Or mAppliedFontBold <> mFontBold _
        Or mAppliedIconGlyph <> mIconGlyph
    textChanged = (Not mAppliedValid) Or (mAppliedText <> shownText)
    If textChanged Then
        ' Replaced text takes its first character's font, which an icon
        ' holds in the icon font.
        If LenB(mAppliedText) = 0 Or LenB(shownText) = 0 _
            Or LenB(mIconGlyph) > 0 Then
            mustRestyle = True
        End If
        mainShape.TextFrame2.TextRange.Text = shownText
        mAppliedText = shownText
    End If
    If mustRestyle Then
        With mainShape.TextFrame2.TextRange.Font
            .Name = themeValue.FontName
            .Size = resolvedSize
            .Bold = BoolToMso(mFontBold Or mKind = KIND_BUTTON Or mKind = KIND_BADGE _
                Or mIsMenu)
            .Fill.ForeColor.RGB = resolvedColor
        End With
        If LenB(mIconGlyph) > 0 Then
            With mainShape.TextFrame2.TextRange.Characters(1, 1).Font
                .Name = ReDimUI.IconFont
                .Bold = msoFalse
            End With
        End If
        mAppliedTextColor = resolvedColor
        mAppliedFontSize = resolvedSize
        mAppliedFontBold = mFontBold
        mAppliedIconGlyph = mIconGlyph
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
            If mSelectOpen And (mKind = KIND_DATE Or mIsMenu) Then CloseDropList
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
            If ClearListFilter() Then
                Exit Sub
            ElseIf mSelectOpen Then
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
        Case KIND_TABS
            HandleTabsKey hostApp, keyText
        Case KIND_DATE
            HandleDateKey hostApp, keyText
        Case KIND_TABLE
            HandleTableKey hostApp, keyText
        Case KIND_EXPANDER
            If keyText = " " Or keyText = "{ENTER}" Then
                ActivateByKey
            ElseIf (keyText = "{RIGHT}" And Not mChecked) _
                Or (keyText = "{LEFT}" And mChecked) Then
                ToggleExpander hostApp
                RunClickHandlers hostApp
            End If
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
' Left and Right move to the tab beside, wrapping at the ends; Home and
' End go to the first and last. The tab shown follows the keys.
Private Sub HandleTabsKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim tabTotal As Long
    Dim aimTab As Long

    tabTotal = ItemCount
    If tabTotal = 0 Then Exit Sub
    Select Case keyText
        Case "{LEFT}"
            aimTab = ShownTab() - 1
            If aimTab < 1 Then aimTab = tabTotal
        Case "{RIGHT}"
            aimTab = ShownTab() + 1
            If aimTab > tabTotal Then aimTab = 1
        Case "{HOME}"
            aimTab = 1
        Case "{END}"
            aimTab = tabTotal
        Case "{ENTER}"
            hostApp.ClickDefaultButton
    End Select
    If aimTab = 0 Then Exit Sub
    If SelectTab(hostApp, aimTab) Then RunClickHandlers hostApp
End Sub

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
        hostApp.SetState mWritesToKey, ItemStateValue(ItemAt(rowNumber))
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
    Dim stepDirection As Long

    itemTotal = ItemCount
    If itemTotal = 0 Then
        If keyText = "{ENTER}" Then hostApp.ClickDefaultButton
        Exit Sub
    End If
    ' Moves land on items that can be picked, passing over disabled
    ' items and group headers in the direction of travel.
    If mSelectOpen Then
        anchorItem = mListHighlight
        If anchorItem < 1 Then anchorItem = CLng(mValue)
        stepDirection = 1
        Select Case keyText
            Case "{UP}"
                aimRow = anchorItem - 1
                stepDirection = -1
            Case "{DOWN}"
                aimRow = anchorItem + 1
            Case "{PGUP}"
                aimRow = anchorItem - mListRows
                stepDirection = -1
            Case "{PGDN}"
                aimRow = anchorItem + mListRows
            Case "{HOME}"
                aimRow = 1
            Case "{END}"
                aimRow = itemTotal
                stepDirection = -1
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
        aimRow = PickableNear(aimRow, stepDirection)
        If aimRow = 0 Then Exit Sub
        HighlightListItem aimRow
        Exit Sub
    End If
    ' A closed menu opens on the keys that open a list; it has no
    ' selection for the arrows to move.
    If mIsMenu Then
        Select Case keyText
            Case " ", "{ENTER}", "{DOWN}", "{UP}", "{ALTDOWN}", "{F4}"
                OpenDropList hostApp.Sheet
                aimRow = PickableNear(1, 1)
                If aimRow > 0 Then HighlightListItem aimRow
        End Select
        Exit Sub
    End If
    anchorItem = CLng(mValue)
    stepDirection = 1
    Select Case keyText
        Case "{UP}", "{LEFT}"
            aimRow = anchorItem - 1
            stepDirection = -1
        Case "{DOWN}", "{RIGHT}"
            aimRow = anchorItem + 1
        Case "{PGUP}"
            aimRow = anchorItem - mListRows
            stepDirection = -1
        Case "{PGDN}"
            aimRow = anchorItem + mListRows
        Case "{HOME}"
            aimRow = 1
        Case "{END}"
            aimRow = itemTotal
            stepDirection = -1
        Case " ", "{ALTDOWN}", "{F4}"
            OpenDropList hostApp.Sheet
            aimRow = PickableNear(IIf(anchorItem >= 1, anchorItem, 1), 1)
            If aimRow > 0 Then HighlightListItem aimRow
            Exit Sub
        Case "{ENTER}"
            hostApp.ClickDefaultButton
            Exit Sub
        Case Else
            aimRow = TypeAheadItem(keyText, anchorItem)
            If aimRow = 0 Then Exit Sub
    End Select
    aimRow = PickableNear(aimRow, stepDirection)
    If aimRow = 0 Then Exit Sub
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
    ' A menu keeps no selection: its pick runs the command.
    If mIsMenu Then
        RunMenuCommand hostApp, ItemAt(itemIndex)
        Exit Function
    End If
    mSelectOpen = False
    mListHighlight = 0
    If itemIndex = CLng(mValue) Then
        Touch
        Exit Function
    End If
    mValue = itemIndex
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, ItemStateValue(ItemAt(itemIndex))
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
        If ItemPickable(candidate) _
            And LCase$(Left$(ItemAt(candidate), Len(typedPrefix))) = typedPrefix Then
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
    Dim shownTotal As Long

    If ItemCount = 0 Then Exit Sub
    If TypeListFilter(keyText) Then Exit Sub
    firstRow = IIf(mSelectAllRow, 0, 1)
    shownTotal = CheckShownCount()
    ' The cursor walks shown rows (0 is select-all); aimRow is a shown
    ' position until the end, where it turns back into an item.
    aimRow = CheckShownPosition(mKeyRow)
    Select Case keyText
        Case "{UP}"
            aimRow = aimRow - 1
        Case "{DOWN}"
            aimRow = aimRow + 1
        Case "{HOME}", "{PGUP}"
            aimRow = firstRow
        Case "{END}", "{PGDN}"
            aimRow = shownTotal
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
    If aimRow > shownTotal Then aimRow = shownTotal
    PlaceKeyCursor 0, IIf(aimRow = 0, 0, CheckShownItem(aimRow))
    Touch
End Sub

' Toggles one row, or with row 0 the select-all header: from mixed or
' none it checks everything, from all it clears. Writes the state and
' moves the keyboard cursor there. True when anything toggled; the
' caller runs the handlers.
' Select-all under a filter: unchecks the rows that show when all of
' them are checked, else checks them; hidden rows keep their checks.
Private Sub ToggleShownChecks()
    Dim idx As Long
    Dim allChecked As Boolean
    Dim anyShown As Boolean

    allChecked = True
    For idx = 1 To ItemCount
        If InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            anyShown = True
            If Not TransferSelectionHas(mCheckedSet, idx) Then allChecked = False
        End If
    Next idx
    If Not anyShown Then Exit Sub
    For idx = 1 To ItemCount
        If InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            If TransferSelectionHas(mCheckedSet, idx) = allChecked Then
                TransferSelectionToggle mCheckedSet, idx
            End If
        End If
    Next idx
End Sub

Private Function ToggleCheckRow(ByVal hostApp As ReDimUI, ByVal checkRow As Long) As Boolean
    Dim rowNumber As Long

    EnsureCheckedSet
    If ItemCount = 0 Then Exit Function
    If checkRow = 0 Then
        If Not mSelectAllRow Then Exit Function
        If LenB(mFilterCheck) > 0 Then
            ToggleShownChecks
        ElseIf mCheckedSet.Count = ItemCount Then
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
    If TypeListFilter(keyText) Then Exit Sub
    rowsShown = TransferRowCapacity()
    ' The cursor walks the rows the panel shows; aimRow is a shown
    ' position until the end, where it turns back into an item.
    aimRow = TransferShownPosition(mKeyPanel, mKeyRow)
    Select Case keyText
        Case "{ALTUP}", "{ALTDOWN}"
            ' Alt+Up and Alt+Down reorder the chosen side from its rows,
            ' and the window follows the cursor's row.
            If Not mReorderable Or mKeyPanel <> 1 Then Exit Sub
            ReorderChosen hostApp, (keyText = "{ALTUP}")
            aimRow = TransferShownPosition(mKeyPanel, mKeyRow)
        Case "{UP}"
            aimRow = aimRow - 1
        Case "{DOWN}"
            aimRow = aimRow + 1
        Case "{HOME}"
            aimRow = 1
        Case "{END}"
            aimRow = TransferShownCount(mKeyPanel)
        Case "{PGUP}"
            aimRow = aimRow - rowsShown
        Case "{PGDN}"
            aimRow = aimRow + rowsShown
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
    rowTotal = TransferShownCount(mKeyPanel)
    If aimRow > rowTotal Then aimRow = rowTotal
    If aimRow < 1 Then aimRow = IIf(rowTotal > 0, 1, 0)
    PlaceKeyCursor mKeyPanel, TransferShownItem(mKeyPanel, aimRow)
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
        DropPart hostSheet, "cx"
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
    If LenB(ShownError()) > 0 And mVisible Then
        noteKey = PartsState(themeValue) & ":" & ShownError()
    End If
    SyncFieldNote hostSheet, themeValue, "me", noteKey, ShownError(), _
        themeValue.DangerColor, msoAlignLeft, noteSize, mAppliedMessageSig
    noteKey = vbNullString
    If mMaxLength > 0 And mVisible Then
        noteKey = PartsState(themeValue) & ":" & Len(mText) & "/" & mMaxLength
    End If
    SyncFieldNote hostSheet, themeValue, "mc", noteKey, _
        Len(mText) & "/" & mMaxLength, themeValue.OnMutedColor, msoAlignRight, _
        noteSize, mAppliedCountSig
End Sub

' The width a badge takes: its text and a little room, never narrower
' than it is tall or than the width it was given.
Private Function BadgeSpan() As Double
    Dim textSize As Double

    textSize = mFontSize
    If textSize <= 0 Then textSize = 9
    BadgeSpan = Len(mText) * textSize * 0.62 + 12
    If BadgeSpan < mHeight Then BadgeSpan = mHeight
    If BadgeSpan < mWidth Then BadgeSpan = mWidth
End Function

' BadgeText's pill: danger-colored on the top-right corner, ringed in the
' canvas color so it reads apart from the control under it.
Private Sub ApplyBadgePart(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim badgeKey As String
    Dim badgeShape As Shape
    Dim badgeSpanNow As Double

    If LenB(mBadgeText) = 0 And LenB(mAppliedBadgeSig) = 0 Then Exit Sub
    If LenB(mBadgeText) > 0 And mVisible Then
        badgeKey = PartsState(themeValue) & ":" & mBadgeText
    End If
    If badgeKey = mAppliedBadgeSig Then Exit Sub
    mAppliedBadgeSig = badgeKey
    If LenB(badgeKey) = 0 Then
        DropPart hostSheet, "bd"
        Exit Sub
    End If
    badgeSpanNow = Len(mBadgeText) * 9 * 0.62 + 10
    If badgeSpanNow < 16 Then badgeSpanNow = 16
    Set badgeShape = EnsurePartShape(hostSheet, "bd", msoShapeRoundedRectangle)
    On Error Resume Next
    badgeShape.Adjustments(1) = 0.5
    On Error GoTo 0
    badgeShape.Left = mLeft + mWidth - badgeSpanNow + 6
    badgeShape.Top = mTop - 7
    badgeShape.Width = badgeSpanNow
    badgeShape.Height = 16
    badgeShape.Fill.Visible = msoTrue
    badgeShape.Fill.ForeColor.RGB = themeValue.DangerColor
    badgeShape.Line.Visible = msoTrue
    badgeShape.Line.ForeColor.RGB = themeValue.CanvasColor
    badgeShape.Line.Weight = 1.5
    With badgeShape.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = mBadgeText
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = 9
            .Bold = msoTrue
            .Fill.ForeColor.RGB = themeValue.OnPrimaryColor
        End With
    End With
    badgeShape.Visible = msoTrue
    badgeShape.OnAction = mAppliedOnAction
End Sub

' The caption above the control and the hint below it, small text; the
' hint gives its place to a float field's message, and leaves room for
' a length count.
Private Sub ApplyAdornments(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim noteSize As Double
    Dim captionKey As String
    Dim hintKey As String

    If LenB(mCaptionText) = 0 And LenB(mAppliedCaptionSig) = 0 _
        And LenB(mHintText) = 0 And LenB(mAppliedHintSig) = 0 Then Exit Sub
    noteSize = mFontSize
    If noteSize <= 0 Then noteSize = themeValue.BaseFontSize
    noteSize = noteSize - 1
    If LenB(mCaptionText) > 0 And mVisible Then
        captionKey = PartsState(themeValue) & ":" & mCaptionText & ":" & CStr(mRequired)
    End If
    SyncFieldCaption hostSheet, themeValue, captionKey, noteSize
    If LenB(mHintText) > 0 And mVisible _
        And Not (IsFloatField() And LenB(ShownError()) > 0) Then
        hintKey = PartsState(themeValue) & ":" & mHintText & ":" & CStr(mMaxLength)
    End If
    SyncFieldNote hostSheet, themeValue, "fh", hintKey, mHintText, _
        themeValue.OnMutedColor, msoAlignLeft, noteSize - 1, mAppliedHintSig
End Sub

Private Sub SyncFieldCaption( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal captionKey As String, _
    ByVal noteSize As Double _
)
    Dim captionShape As Shape
    Dim captionWords As String

    If captionKey = mAppliedCaptionSig Then Exit Sub
    mAppliedCaptionSig = captionKey
    If LenB(captionKey) = 0 Then
        DropPart hostSheet, "fc"
        Exit Sub
    End If
    captionWords = mCaptionText
    If mRequired Then captionWords = captionWords & " *"
    Set captionShape = EnsurePartShape(hostSheet, "fc", msoShapeRectangle)
    captionShape.Left = mLeft
    captionShape.Top = mTop - noteSize - 8
    If captionShape.Top < 0 Then captionShape.Top = 0
    captionShape.Width = mWidth
    captionShape.Height = noteSize + 7
    captionShape.Fill.Visible = msoFalse
    captionShape.Line.Visible = msoFalse
    With captionShape.TextFrame2
        .MarginLeft = 1
        .MarginRight = 1
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorBottom
        .WordWrap = msoFalse
        .TextRange.Text = captionWords
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = noteSize
            .Bold = msoFalse
            .Fill.ForeColor.RGB = ResolveDrawnInk(themeValue)
        End With
        If mRequired Then
            .TextRange.Characters(Len(captionWords), 1).Font.Fill.ForeColor.RGB = _
                themeValue.DangerColor
        End If
    End With
    captionShape.Visible = msoTrue
    captionShape.OnAction = mAppliedOnAction
End Sub

' The Skeleton's muted fill at its pulse step: toward the surface and
' back over 1.4 seconds, in 16 steps; flat when motion is reduced.
Private Function SkeletonFill(ByVal themeValue As ReDimUI) As Long
    Dim stepNow As Long

    stepNow = PulseStep()
    mAppliedPulseStep = stepNow
    SkeletonFill = BlendRgb(themeValue.MutedColor, themeValue.SurfaceColor, _
        0.5 * (1 - Cos(6.2831853 * stepNow / 16)) * 0.6)
End Function

Private Function PulseStep() As Long
    If ReDimUI.MotionReduced Then Exit Function
    PulseStep = Int(((mPulseMs Mod 1400) / 1400) * 16)
End Function

' A Skeleton's text bars: evenly spaced, the last at 60 percent.
Private Sub ApplySkeletonBars(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim skeletonBar As Shape
    Dim barNo As Long
    Dim barPitch As Double
    Dim barFill As Long

    barFill = SkeletonFill(themeValue)
    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mSkeletonLines) & _
        ":" & CStr(barFill)) Then Exit Sub
    If mSkeletonLines > 0 Then barPitch = mHeight / mSkeletonLines
    For barNo = 1 To mSkeletonLines
        Set skeletonBar = EnsurePartShape(hostSheet, "sk" & barNo, msoShapeRoundedRectangle)
        skeletonBar.Left = mLeft
        skeletonBar.Top = mTop + (barNo - 1) * barPitch + barPitch * 0.2
        skeletonBar.Height = barPitch * 0.6
        skeletonBar.Width = mWidth * IIf(barNo = mSkeletonLines And barNo > 1, 0.6, 1)
        skeletonBar.Fill.Visible = msoTrue
        skeletonBar.Fill.ForeColor.RGB = barFill
        skeletonBar.Line.Visible = msoFalse
        skeletonBar.Visible = BoolToMso(mVisible)
        skeletonBar.OnAction = mAppliedOnAction
    Next barNo
    ' Bars past the count are leftovers of a longer skeleton.
    For barNo = mSkeletonLines + 1 To mAppliedOptionCount
        DropPart hostSheet, "sk" & barNo
    Next barNo
    mAppliedOptionCount = mSkeletonLines
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
        DropPart hostSheet, partName
        Exit Sub
    End If
    appliedKey = noteKey
    Set noteShape = EnsurePartShape(hostSheet, partName, msoShapeRectangle)
    noteShape.Left = mLeft
    noteShape.Top = mTop + mHeight + 1
    noteShape.Width = mWidth
    If partName = "fh" And mMaxLength > 0 Then noteShape.Width = mWidth - 44
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

' The message a float field shows: its validation's, else the one set
' from outside.
Private Function ShownError() As String
    If LenB(mValidationError) > 0 Then
        ShownError = mValidationError
    Else
        ShownError = mErrorText
    End If
End Function

' ValidateAll's fields: an enabled float field with Required or
' Validates that shows, or would once its panels show.
Friend Function TakesChecks(ByVal hostApp As ReDimUI) As Boolean
    If Not IsFloatField() Or Not mEnabled Then Exit Function
    If Not mRequired And LenB(mValidateProc) = 0 Then Exit Function
    TakesChecks = Revealable(hostApp)
End Function

' Runs the commit's check on the text the field holds and shows the
' outcome. True when it passes.
Friend Function PassesChecks(ByVal hostApp As ReDimUI) As Boolean
    RunValidation hostApp
    Touch
    PassesChecks = (LenB(mValidationError) = 0)
End Function

' Whether the control shows, or would once each Tabs or Expander panel
' it sits on shows: False when it, or a stack or panel control holding
' it, is hidden in its own right.
Friend Function Revealable(ByVal hostApp As ReDimUI) As Boolean
    Dim holderValue As ReDimUI

    If Not mTabShut Then
        Revealable = mVisible
        Exit Function
    End If
    If Not mVisibleBeforeTab Then Exit Function
    If LenB(mStackId) > 0 Then
        Set holderValue = hostApp.PeekComponent(mStackId)
        If Not holderValue Is Nothing Then
            If Not holderValue.Revealable(hostApp) Then Exit Function
        End If
    End If
    If LenB(mTabOwnerId) > 0 Then
        Set holderValue = hostApp.PeekComponent(mTabOwnerId)
        If Not holderValue Is Nothing Then
            If Not holderValue.Revealable(hostApp) Then Exit Function
        End If
    End If
    Revealable = True
End Function

' Shows the panels the control sits on, outermost first: a Tabs control
' turns to its tab and an Expander opens, each writing its WritesTo
' state and firing nothing, as from code.
Friend Sub RevealPanels(ByVal hostApp As ReDimUI)
    Dim holderValue As ReDimUI

    If LenB(mStackId) > 0 Then
        Set holderValue = hostApp.PeekComponent(mStackId)
        If Not holderValue Is Nothing Then
            holderValue.RevealPanels hostApp
        End If
    End If
    If LenB(mTabOwnerId) = 0 Then Exit Sub
    Set holderValue = hostApp.PeekComponent(mTabOwnerId)
    If holderValue Is Nothing Then Exit Sub
    holderValue.RevealPanels hostApp
    holderValue.ShowPanel hostApp, mTabNumber
End Sub

Friend Sub ShowPanel(ByVal hostApp As ReDimUI, ByVal tabNumber As Long)
    If mKind = KIND_TABS Then
        SelectTab hostApp, tabNumber
    ElseIf mKind = KIND_EXPANDER Then
        If mChecked Then Exit Sub
        mChecked = True
        If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, True
        Touch
    End If
End Sub

' Runs Validates' check on the text. A check that fails to run counts
' as a pass and goes to the app's error handler.
Private Sub RunValidation(ByVal hostApp As ReDimUI)
    Dim verdict As Variant

    ' Required speaks first: an empty text never reaches the check.
    If mRequired And LenB(Trim$(mText)) = 0 Then
        If LenB(mRequiredMessage) > 0 Then
            mValidationError = mRequiredMessage
        Else
            mValidationError = Phrase(UI_REQUIRED)
        End If
        mValidationLive = True
        Exit Sub
    End If
    If LenB(mValidateProc) = 0 Then
        mValidationError = vbNullString
        mValidationLive = False
        Exit Sub
    End If
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
        DropPart hostSheet, "fr"
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
        Set cursorShape = EnsurePartShape(hostSheet, "kr", msoShapeRoundedRectangle)
        cursorShape.Left = cursorLeft
        cursorShape.Top = cursorTop
        cursorShape.Width = cursorWidth
        cursorShape.Height = cursorHeight
        RoundCorners cursorShape, cursorHeight
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
        DropPart hostSheet, "kr"
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
        Case KIND_TABS
            ApplyTabs mainShape, hostSheet, themeValue
        Case KIND_DATE
            ApplyListCaret hostSheet, themeValue, 14
            SyncCalendar hostSheet, themeValue
        Case KIND_TABLE
            ApplyTable hostSheet, themeValue
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
        Case KIND_SKELETON
            ApplySkeletonBars hostSheet, themeValue
        Case KIND_SPARKLINE
            ApplySparkline hostSheet, themeValue
    End Select
End Sub

' What a screen reader hears for a sparkline: how many values, the low
' and the high, and the last.
Private Function SparklineSummary() As String
    Dim lowValue As Double
    Dim highValue As Double

    If mSparkCount = 0 Then
        SparklineSummary = Phrase(UI_ALT_SPARKLINE_EMPTY)
        Exit Function
    End If
    SparkRange lowValue, highValue
    SparklineSummary = Phrase(UI_ALT_SPARKLINE, mSparkCount, lowValue, highValue, _
        mSparkValues(mSparkCount))
End Function

' The lowest and highest of a sparkline's values; it has at least one.
Private Sub SparkRange(ByRef lowValue As Double, ByRef highValue As Double)
    Dim pointNo As Long

    lowValue = mSparkValues(1)
    highValue = mSparkValues(1)
    For pointNo = 2 To mSparkCount
        If mSparkValues(pointNo) < lowValue Then lowValue = mSparkValues(pointNo)
        If mSparkValues(pointNo) > highValue Then highValue = mSparkValues(pointNo)
    Next pointNo
End Sub

' A sparkline's line, one polyline through its values scaled to its
' height, and a dot on the last value, drawn again only when the values,
' the place, the color, or the visibility changed. Equal values run
' along the middle; one value is a dot alone.
Private Sub ApplySparkline(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim inkRgb As Long
    Dim sparkKey As String
    Dim lowValue As Double
    Dim highValue As Double
    Dim pointNo As Long
    Dim coords() As Single
    Dim lineShape As Shape
    Dim dotShape As Shape

    inkRgb = ResolveFill(themeValue)
    sparkKey = PartsState(themeValue) & ":" & CStr(mSparkVersion) & ":" & CStr(inkRgb)
    If sparkKey = mAppliedSparkSig Then Exit Sub
    mAppliedSparkSig = sparkKey
    DropPart hostSheet, "sl"
    If mSparkCount = 0 Then
        DropPart hostSheet, "sd"
        Exit Sub
    End If
    SparkRange lowValue, highValue
    ReDim coords(1 To mSparkCount, 1 To 2)
    For pointNo = 1 To mSparkCount
        If mSparkCount = 1 Then
            coords(pointNo, 1) = mLeft + mWidth / 2
        Else
            coords(pointNo, 1) = mLeft + SPARK_DOT / 2 + _
                (mWidth - SPARK_DOT) * (pointNo - 1) / (mSparkCount - 1)
        End If
        If highValue > lowValue Then
            coords(pointNo, 2) = mTop + SPARK_DOT / 2 + (mHeight - SPARK_DOT) * _
                (highValue - mSparkValues(pointNo)) / (highValue - lowValue)
        Else
            coords(pointNo, 2) = mTop + mHeight / 2
        End If
    Next pointNo
    If mSparkCount > 1 Then
        Set lineShape = hostSheet.Shapes.AddPolyline(coords)
        lineShape.Name = ShapeName & PART_SEPARATOR & "sl"
        OwnerApp.NoteShapeDrawn lineShape.Name
        lineShape.Placement = xlFreeFloating
        lineShape.Fill.Visible = msoFalse
        lineShape.Line.Visible = msoTrue
        lineShape.Line.ForeColor.RGB = inkRgb
        lineShape.Line.Weight = 1.5
        lineShape.OnAction = mAppliedOnAction
        lineShape.Visible = BoolToMso(mVisible)
    End If
    Set dotShape = EnsurePartShape(hostSheet, "sd", msoShapeOval)
    dotShape.Left = coords(mSparkCount, 1) - SPARK_DOT / 2
    dotShape.Top = coords(mSparkCount, 2) - SPARK_DOT / 2
    dotShape.Width = SPARK_DOT
    dotShape.Height = SPARK_DOT
    dotShape.Fill.Visible = msoTrue
    dotShape.Fill.ForeColor.RGB = inkRgb
    dotShape.OnAction = mAppliedOnAction
    dotShape.Visible = BoolToMso(mVisible)
    ' The line was made after the dot, so the dot comes back on top.
    If mSparkCount > 1 Then dotShape.ZOrder msoBringToFront
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

' The card fitted to its message, the tone's icon color and edge, the
' close button, and the action button, placed at the toast's current
' visual top; the pump carries them through the slide and the fade from
' there. The card fits first, so the buttons sit on its final edges.
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
    If ExtrasCurrent(ToastPartsKey(themeValue, visualTop)) Then Exit Sub
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    ' Margins keep the words clear of the buttons; the icon's tab stop
    ' is set before the fit, since it moves where the words begin.
    mainShape.TextFrame2.MarginLeft = 8
    If LenB(mToastActionLabel) > 0 Then
        ' Sized from the label; ink in surface text, since the accent
        ' does not reach 4.5:1 as text on every theme's surface.
        actionWidth = Len(mToastActionLabel) * resolvedSize * 0.62 + 14
        mainShape.TextFrame2.MarginRight = PAGER_TARGET + 11 + actionWidth
    Else
        mainShape.TextFrame2.MarginRight = PAGER_TARGET + 6
    End If
    If LenB(ToastIcon()) > 0 Then
        With mainShape.TextFrame2.TextRange.ParagraphFormat.TabStops
            If .Count = 0 Then .Add msoTabStopLeft, 18
        End With
    End If
    FitToastCard mainShape
    If LenB(ToastIcon()) > 0 Then
        mainShape.Line.ForeColor.RGB = ToastToneColor(themeValue)
        mainShape.Line.Weight = 1.5
        mainShape.TextFrame2.TextRange.Characters(1, 1).Font.Fill.ForeColor.RGB = _
            ToastToneColor(themeValue)
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
        Set actionButton = EnsurePartShape( _
            hostSheet, "ta", msoShapeRoundedRectangle)
        actionButton.Left = mLeft + mWidth - PAGER_TARGET - 7 - actionWidth
        actionButton.Top = visualTop + (mHeight - 22) / 2
        actionButton.Width = actionWidth
        actionButton.Height = 22
        actionButton.Fill.Visible = msoTrue
        actionButton.Fill.ForeColor.RGB = ResolveSurfaceFill(themeValue)
        actionButton.Line.Visible = msoTrue
        If mToastActionBorder >= 0 Then
            actionButton.Line.ForeColor.RGB = mToastActionBorder
        Else
            actionButton.Line.ForeColor.RGB = themeValue.PrimaryColor
        End If
        actionButton.Line.Weight = 1
        StyleGlyphText actionButton, mToastActionLabel, _
            themeValue.OnSurfaceColor, resolvedSize - 1, themeValue
        actionButton.Visible = BoolToMso(mVisible)
        actionButton.OnAction = mAppliedOnAction
    End If
    ' The fit may have resized the card: record the state the parts now
    ' match, so the next apply skips them.
    mAppliedExtrasSig = ToastPartsKey(themeValue, visualTop)
End Sub

Private Function ToastPartsKey(ByVal themeValue As ReDimUI, ByVal visualTop As Double) As String
    ToastPartsKey = PartsState(themeValue) & ":" & ToastIcon() & ":" & mToastActionLabel & _
        ":" & CStr(mToastActionBorder) & ":" & CStr(visualTop) & ":" & CStr(mToastMinWidth) & _
        ":" & CStr(mToastMaxWidth) & ":" & CStr(mToastMaxHeight) & Chr$(2) & ResolveText(Nothing)
End Function

' The card fits its message: as wide as the words need between MinWidth
' and MaxWidth, and no wider than the window's room to the right, then
' as tall as the words wrapped at that width, from one line's height to
' MaxHeight, where a longer message ends in an ellipsis. A change of
' height restacks the tray once the flush ends.
Private Sub FitToastCard(ByVal cardShape As Shape)
    Dim minWide As Double
    Dim maxWide As Double
    Dim maxTall As Double
    Dim fitWidth As Double
    Dim fitHeight As Double
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double
    Dim priorHeight As Double

    minWide = mToastMinWidth
    If minWide <= 0 Then minWide = TOAST_WIDTH
    maxWide = mToastMaxWidth
    If maxWide <= 0 Then maxWide = TOAST_WIDTH
    OwnerApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    If mLeft + maxWide > viewLeft + viewWidth - 12 Then
        maxWide = viewLeft + viewWidth - 12 - mLeft
    End If
    If maxWide < minWide Then maxWide = minWide
    maxTall = mToastMaxHeight
    If maxTall <= 0 Then maxTall = TOAST_MAX_HEIGHT
    If maxTall < TOAST_MIN_HEIGHT Then maxTall = TOAST_MIN_HEIGHT
    ' A rounded card lays its words out inside a frame inset by its corner
    ' radius times 0.2929 (Office's preset geometry), and the radius is a
    ' share of the shorter side, so the corner is held at TOAST_CORNER
    ' points: the inset stays put however tall the card grows.
    With cardShape.TextFrame2
        .WordWrap = msoFalse
        fitWidth = .TextRange.BoundWidth + .MarginLeft + .MarginRight + TOAST_TEXT_INSET * 2 + 1
        If fitWidth < minWide Then fitWidth = minWide
        If fitWidth > maxWide Then fitWidth = maxWide
        cardShape.Width = fitWidth
        HoldToastCorner cardShape
        .WordWrap = msoTrue
        fitHeight = .TextRange.BoundHeight + .MarginTop + .MarginBottom + TOAST_TEXT_INSET * 2 + 6
        If fitHeight > maxTall Then
            ShortenToastText cardShape, maxTall
            fitHeight = maxTall
        End If
    End With
    If fitHeight < TOAST_MIN_HEIGHT Then fitHeight = TOAST_MIN_HEIGHT
    priorHeight = mHeight
    mWidth = fitWidth
    mHeight = fitHeight
    cardShape.Height = fitHeight
    HoldToastCorner cardShape
    If Abs(fitHeight - priorHeight) > 0.01 Then OwnerApp.NoteToastResized
End Sub

Private Sub HoldToastCorner(ByVal cardShape As Shape)
    Dim shortSide As Double

    shortSide = cardShape.Width
    If cardShape.Height < shortSide Then shortSide = cardShape.Height
    If shortSide <= 0 Then Exit Sub
    On Error Resume Next
    cardShape.Adjustments(1) = TOAST_CORNER / shortSide
    On Error GoTo 0
End Sub

' Cuts a message that runs past MaxHeight back to the lines that fit,
' the last word giving way to an ellipsis. The words keep their ink.
Private Sub ShortenToastText(ByVal cardShape As Shape, ByVal maxTall As Double)
    Dim lineTotal As Long
    Dim fitLines As Long
    Dim keptWords As String
    Dim cutAt As Long
    Dim inkRgb As Long

    With cardShape.TextFrame2
        lineTotal = .TextRange.Lines.Count
        If lineTotal < 2 Then Exit Sub
        fitLines = Int((maxTall - .MarginTop - .MarginBottom - TOAST_TEXT_INSET * 2 - 6) / _
            (.TextRange.BoundHeight / lineTotal))
        If fitLines < 1 Then fitLines = 1
        If fitLines >= lineTotal Then Exit Sub
        keptWords = .TextRange.Text
        With .TextRange.Lines(fitLines)
            keptWords = RTrim$(Left$(keptWords, .Start + .Length - 1))
        End With
        cutAt = InStrRev(keptWords, " ")
        If cutAt > Len(keptWords) \ 2 Then keptWords = RTrim$(Left$(keptWords, cutAt - 1))
        inkRgb = .TextRange.Characters(.TextRange.Length, 1).Font.Fill.ForeColor.RGB
        .TextRange.Text = keptWords & ChrW(8230)
        .TextRange.Font.Fill.ForeColor.RGB = inkRgb
    End With
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
    Erase mAppliedTableRowStates
    Erase mAppliedTableHeadStates
    mAppliedTableFootSig = "?"
    mAppliedSparkSig = vbNullString
    mGlidePart = vbNullString
    mAppliedIndicatorTab = 0
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
        ":" & mText & ":" & mAccessKey & ":" & PointerKey()) Then Exit Sub
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
        mainShape.Fill.ForeColor.RGB = PointerTint(ResolveFill(themeValue), _
            themeValue.OnPrimaryColor)
        mainShape.Line.Visible = msoFalse
    Else
        mainShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        mainShape.Line.Visible = msoTrue
        mainShape.Line.ForeColor.RGB = _
            IIf(PointerOn(), themeValue.PrimaryColor, themeValue.BorderColor)
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
' A Tabs control: a rule along its foot, one tab per item sized to its
' text (equal shares when the texts need more than the width, cut short
' with an ellipsis where they do not fit), and an accent bar under the
' tab shown.
Private Sub ApplyTabs( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim tabTotal As Long
    Dim tabNow As Long
    Dim textSize As Double
    Dim tabWidths() As Double
    Dim widthSum As Double
    Dim tabLeft As Double
    Dim idx As Long
    Dim tabShape As Shape
    Dim accentBar As Shape
    Dim tabFill As Long
    Dim tabWords As String
    Dim fitChars As Long
    Dim barMade As Boolean
    Dim barWidth As Double
    Dim glideBase As String
    Dim tabMade As Boolean

    tabTotal = ItemCount
    tabNow = ShownTab()
    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(tabNow) & ":" & _
        PointerKey() & ":" & CStr(mItemsVersion)) Then Exit Sub
    mainShape.Left = mLeft
    mainShape.Top = mTop + mHeight - 1
    mainShape.Width = mWidth
    mainShape.Height = 1
    mainShape.Fill.Visible = msoTrue
    mainShape.Fill.ForeColor.RGB = themeValue.BorderColor
    mainShape.Line.Visible = msoFalse
    mainShape.Visible = BoolToMso(mVisible)
    textSize = mFontSize
    If textSize <= 0 Then textSize = themeValue.BaseFontSize
    If tabTotal > 0 Then
        EnsureItemTexts
        ReDim tabWidths(1 To tabTotal)
        ReDim mTabRights(1 To tabTotal)
        For idx = 1 To tabTotal
            tabWidths(idx) = Len(mItemTexts(idx)) * textSize * 0.6 + 24
            widthSum = widthSum + tabWidths(idx)
        Next idx
        If widthSum > mWidth Then
            For idx = 1 To tabTotal
                tabWidths(idx) = mWidth / tabTotal
            Next idx
        End If
    End If
    tabLeft = mLeft
    For idx = 1 To tabTotal
        Set tabShape = EnsurePartShape(hostSheet, "tb" & idx, msoShapeRoundedRectangle, tabMade)
        tabShape.Left = tabLeft
        tabShape.Top = mTop
        tabShape.Width = tabWidths(idx)
        ' The foot stays clear for the rule and the accent bar.
        tabShape.Height = mHeight - 3
        If tabMade Or Not mAppliedValid Then RoundCorners tabShape, mHeight - 3
        tabFill = PointerTint(themeValue.SurfaceColor, themeValue.OnSurfaceColor, _
            "tab" & idx)
        tabShape.Fill.Visible = BoolToMso(tabFill <> themeValue.SurfaceColor)
        If tabFill <> themeValue.SurfaceColor Then tabShape.Fill.ForeColor.RGB = tabFill
        tabShape.Line.Visible = msoFalse
        tabWords = mItemTexts(idx)
        fitChars = Int((tabWidths(idx) - 8) / (textSize * 0.6))
        If Len(tabWords) > fitChars And fitChars >= 2 Then
            tabWords = Left$(tabWords, fitChars - 1) & ChrW(8230)
        End If
        With tabShape.TextFrame2
            .MarginLeft = 4
            .MarginRight = 4
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .TextRange.Text = tabWords
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            With .TextRange.Font
                .Name = themeValue.FontName
                .Size = textSize
                .Bold = BoolToMso(idx = tabNow)
                If idx = tabNow And mEnabled Then
                    .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
                Else
                    .Fill.ForeColor.RGB = themeValue.OnMutedColor
                End If
            End With
        End With
        tabShape.Visible = BoolToMso(mVisible)
        tabShape.OnAction = mAppliedOnAction
        tabLeft = tabLeft + tabWidths(idx)
        mTabRights(idx) = tabLeft - mLeft
    Next idx
    If tabNow >= 1 Then
        Set accentBar = EnsurePartShape(hostSheet, "ti", msoShapeRoundedRectangle, barMade)
        barWidth = tabWidths(tabNow) - 16
        If barWidth < 8 Then barWidth = 8
        ' Another tab shown on a strip drawn in place already glides the
        ' bar over to it.
        glideBase = PartsState(themeValue) & ":" & CStr(mItemsVersion)
        PlaceGliding accentBar, "ti", _
            mLeft + mTabRights(tabNow) - tabWidths(tabNow) + (tabWidths(tabNow) - barWidth) / 2, _
            barWidth, Not barMade And mAppliedValid And mAppliedIndicatorTab >= 1 _
            And mAppliedIndicatorTab <> tabNow And glideBase = mAppliedGlideBase
        mAppliedGlideBase = glideBase
        mAppliedIndicatorTab = tabNow
        accentBar.Top = mTop + mHeight - 3
        accentBar.Height = 3
        accentBar.Fill.Visible = msoTrue
        If mEnabled Then
            accentBar.Fill.ForeColor.RGB = themeValue.PrimaryColor
        Else
            accentBar.Fill.ForeColor.RGB = themeValue.OnMutedColor
        End If
        accentBar.Line.Visible = msoFalse
        accentBar.Visible = BoolToMso(mVisible)
        accentBar.OnAction = mAppliedOnAction
    Else
        DropPart hostSheet, "ti"
        If mGlidePart = "ti" Then mGlidePart = vbNullString
        mAppliedIndicatorTab = 0
    End If
    ' Tabs past the count are leftovers of a longer strip.
    For idx = tabTotal + 1 To mAppliedOptionCount
        DropPart hostSheet, "tb" & idx
    Next idx
    mAppliedOptionCount = tabTotal
End Sub

' Sets the column count, keeping the cells, headers, formats, and widths
' of the columns that remain.
Private Sub ResizeTableColumns(ByVal columnTotal As Long)
    Dim keptCells() As Variant
    Dim rowNo As Long
    Dim colNo As Long

    If columnTotal < 0 Then columnTotal = 0
    If columnTotal = mTableCols Then Exit Sub
    If columnTotal = 0 Then
        mTableCols = 0
        mTableRows = 0
        Erase mTableCells
        Exit Sub
    End If
    If mTableRows > 0 Then
        ReDim keptCells(1 To columnTotal, 1 To mTableRows)
        For rowNo = 1 To mTableRows
            For colNo = 1 To columnTotal
                If colNo <= mTableCols Then keptCells(colNo, rowNo) = mTableCells(colNo, rowNo)
            Next colNo
        Next rowNo
        mTableCells = keptCells
    Else
        ReDim mTableCells(1 To columnTotal, 1 To 4)
    End If
    If mTableCols = 0 Then
        ReDim mTableHeads(1 To columnTotal)
        ReDim mColumnFormats(1 To columnTotal)
        ReDim mColumnPoints(1 To columnTotal)
    Else
        ReDim Preserve mTableHeads(1 To columnTotal)
        ReDim Preserve mColumnFormats(1 To columnTotal)
        ReDim Preserve mColumnPoints(1 To columnTotal)
    End If
    mTableCols = columnTotal
    If mSortColumn > mTableCols Then mSortColumn = 0
End Sub

' Room for at least rowsWanted rows, doubling so appends stay cheap.
Private Sub EnsureTableRoom(ByVal rowsWanted As Long)
    Dim rowsHeld As Long

    rowsHeld = UBound(mTableCells, 2)
    If rowsHeld >= rowsWanted Then Exit Sub
    Do While rowsHeld < rowsWanted
        rowsHeld = rowsHeld * 2
    Loop
    ReDim Preserve mTableCells(1 To mTableCols, 1 To rowsHeld)
End Sub

' After the data changed: the order, the layout, and the numeric columns
' are read again, and a selection past the rows goes.
Private Sub TableChanged()
    mTableVersion = mTableVersion + 1
    mTableOrderCurrent = False
    mTableShownKey = vbNullString
    mColumnLayoutKey = vbNullString
    If CLng(mValue) > mTableRows Then mValue = 0
    Touch
End Sub

' The rows in the order they show: as they came in, or stably sorted by
' the sort column so rows with equal cells keep their order.
Private Sub EnsureTableOrder()
    Dim rowNo As Long
    Dim scratch() As Long

    If mTableOrderCurrent Then Exit Sub
    mTableOrderCurrent = True
    mTableOrderGen = mTableOrderGen + 1
    If mTableRows = 0 Then Exit Sub
    ReDim mTableOrder(1 To mTableRows)
    For rowNo = 1 To mTableRows
        mTableOrder(rowNo) = rowNo
    Next rowNo
    If mSortColumn = 0 Or mTableRows < 2 Then Exit Sub
    ReDim scratch(1 To mTableRows)
    MergeTableRows 1, mTableRows, scratch
End Sub

Private Sub MergeTableRows(ByVal lowAt As Long, ByVal highAt As Long, ByRef scratch() As Long)
    Dim midAt As Long
    Dim leftAt As Long
    Dim rightAt As Long
    Dim outAt As Long

    If highAt <= lowAt Then Exit Sub
    midAt = (lowAt + highAt) \ 2
    MergeTableRows lowAt, midAt, scratch
    MergeTableRows midAt + 1, highAt, scratch
    leftAt = lowAt
    rightAt = midAt + 1
    outAt = lowAt
    Do While leftAt <= midAt And rightAt <= highAt
        If CompareTableRows(mTableOrder(rightAt), mTableOrder(leftAt)) < 0 Then
            scratch(outAt) = mTableOrder(rightAt)
            rightAt = rightAt + 1
        Else
            scratch(outAt) = mTableOrder(leftAt)
            leftAt = leftAt + 1
        End If
        outAt = outAt + 1
    Loop
    Do While leftAt <= midAt
        scratch(outAt) = mTableOrder(leftAt)
        leftAt = leftAt + 1
        outAt = outAt + 1
    Loop
    Do While rightAt <= highAt
        scratch(outAt) = mTableOrder(rightAt)
        rightAt = rightAt + 1
        outAt = outAt + 1
    Loop
    For outAt = lowAt To highAt
        mTableOrder(outAt) = scratch(outAt)
    Next outAt
End Sub

' The rows the table shows: its order, less the rows the filter hides.
' Rebuilt only when the filter, the order, or the data changed.
Private Sub EnsureTableShown()
    Dim shownKey As String
    Dim displayNo As Long
    Dim sourceRow As Long
    Dim colNo As Long

    EnsureTableOrder
    shownKey = mTableFilter & Chr$(1) & CStr(mTableOrderGen) & Chr$(1) & CStr(mTableVersion)
    If shownKey = mTableShownKey Then Exit Sub
    mTableShownKey = shownKey
    mTableShownCount = 0
    If mTableRows = 0 Then Exit Sub
    ReDim mTableShown(1 To mTableRows)
    For displayNo = 1 To mTableRows
        sourceRow = mTableOrder(displayNo)
        If LenB(mTableFilter) = 0 Then
            mTableShownCount = mTableShownCount + 1
            mTableShown(mTableShownCount) = sourceRow
        Else
            For colNo = 1 To mTableCols
                If InStr(1, TableCellText(colNo, mTableCells(colNo, sourceRow)), _
                    mTableFilter, vbTextCompare) > 0 Then
                    mTableShownCount = mTableShownCount + 1
                    mTableShown(mTableShownCount) = sourceRow
                    Exit For
                End If
            Next colNo
        End If
    Next displayNo
End Sub

Private Sub SetTableFilter(ByVal filterText As String)
    If filterText = mTableFilter Then Exit Sub
    mTableFilter = filterText
    mTableOffset = 0
    Touch
End Sub

' Typing while the table has the keys builds its filter; Backspace
' takes a letter off.
Private Function TypeTableFilter(ByVal keyText As String) As Boolean
    Select Case keyText
        Case "{APOS}"
            keyText = "'"
        Case "{QUOTE}"
            keyText = """"
    End Select
    If keyText = "{BS}" Then
        If LenB(mTableFilter) = 0 Then Exit Function
        SetTableFilter Left$(mTableFilter, Len(mTableFilter) - 1)
    ElseIf Len(keyText) = 1 Then
        SetTableFilter mTableFilter & keyText
    Else
        Exit Function
    End If
    TypeTableFilter = True
End Function

' Where a source row sits among the rows shown; 0 when the filter hides
' it or it is no row.
Private Function TableShownAt(ByVal sourceRow As Long) As Long
    Dim displayNo As Long

    If sourceRow < 1 Then Exit Function
    EnsureTableShown
    For displayNo = 1 To mTableShownCount
        If mTableShown(displayNo) = sourceRow Then
            TableShownAt = displayNo
            Exit Function
        End If
    Next displayNo
End Function

' Ctrl+C: the selected row, or every row shown when none is selected or
' the filter hides it, under the header row, tab-separated for pasting
' into the sheet. Dates go as yyyy-mm-dd so the sheet reads them back
' as dates.
Private Sub CopyTableRows()
    Dim clipText As String
    Dim colNo As Long
    Dim displayNo As Long

    If mTableCols = 0 Then Exit Sub
    EnsureTableShown
    For colNo = 1 To mTableCols
        clipText = clipText & IIf(colNo > 1, vbTab, vbNullString) & ClipCell(mTableHeads(colNo))
    Next colNo
    If TableShownAt(CLng(mValue)) > 0 Then
        clipText = clipText & vbCrLf & ClipRow(CLng(mValue))
    Else
        For displayNo = 1 To mTableShownCount
            clipText = clipText & vbCrLf & ClipRow(mTableShown(displayNo))
        Next displayNo
    End If
    WriteClipboardText clipText
End Sub

Private Function ClipRow(ByVal sourceRow As Long) As String
    Dim colNo As Long

    For colNo = 1 To mTableCols
        If colNo > 1 Then ClipRow = ClipRow & vbTab
        ClipRow = ClipRow & ClipCell(mTableCells(colNo, sourceRow))
    Next colNo
End Function

Private Function ClipCell(ByVal cellEntry As Variant) As String
    If IsError(cellEntry) Then
        ClipCell = ErrorCellText(cellEntry)
    ElseIf IsEmpty(cellEntry) Or IsNull(cellEntry) Then
        Exit Function
    ElseIf VarType(cellEntry) = vbDate Then
        If cellEntry = Int(cellEntry) Then
            ClipCell = Format$(cellEntry, "yyyy-mm-dd")
        Else
            ClipCell = Format$(cellEntry, "yyyy-mm-dd hh:nn:ss")
        End If
    Else
        ClipCell = Replace(Replace(Replace(CStr(cellEntry), vbCr, " "), vbLf, " "), _
            vbTab, " ")
    End If
End Function

' Orders two rows by the sort column: numbers and dates by value, before
' text, text without case, and empty cells last whichever the direction.
Private Function CompareTableRows(ByVal rowA As Long, ByVal rowB As Long) As Long
    Dim cellA As Variant
    Dim cellB As Variant
    Dim blankA As Boolean
    Dim blankB As Boolean

    cellA = mTableCells(mSortColumn, rowA)
    cellB = mTableCells(mSortColumn, rowB)
    blankA = IsBlankCell(cellA)
    blankB = IsBlankCell(cellB)
    If blankA Or blankB Then
        If blankA And Not blankB Then CompareTableRows = 1
        If blankB And Not blankA Then CompareTableRows = -1
        Exit Function
    End If
    If IsNumberCell(cellA) And IsNumberCell(cellB) Then
        If CDbl(cellA) < CDbl(cellB) Then
            CompareTableRows = -1
        ElseIf CDbl(cellA) > CDbl(cellB) Then
            CompareTableRows = 1
        End If
    ElseIf IsNumberCell(cellA) Then
        CompareTableRows = -1
    ElseIf IsNumberCell(cellB) Then
        CompareTableRows = 1
    Else
        CompareTableRows = StrComp(TableCellText(0, cellA), TableCellText(0, cellB), _
            vbTextCompare)
    End If
    If mSortDescending Then CompareTableRows = -CompareTableRows
End Function

Private Function IsBlankCell(ByVal cellEntry As Variant) As Boolean
    If IsEmpty(cellEntry) Or IsNull(cellEntry) Then
        IsBlankCell = True
    ElseIf VarType(cellEntry) = vbString Then
        IsBlankCell = (LenB(cellEntry) = 0)
    End If
End Function

Private Function IsNumberCell(ByVal cellEntry As Variant) As Boolean
    Select Case VarType(cellEntry)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, _
            vbDate, vbByte, vbLongLong
            IsNumberCell = True
    End Select
End Function

' A cell as the table shows it: through its column's Format$ pattern
' when it has one (column 0 for none), and Excel's own words for an
' error value.
Private Function TableCellText(ByVal colNo As Long, ByVal cellEntry As Variant) As String
    If IsError(cellEntry) Then
        TableCellText = ErrorCellText(cellEntry)
    ElseIf IsEmpty(cellEntry) Or IsNull(cellEntry) Then
        Exit Function
    ElseIf colNo >= 1 And colNo <= mTableCols Then
        If LenB(mColumnFormats(colNo)) > 0 Then
            TableCellText = Format$(cellEntry, mColumnFormats(colNo))
        Else
            TableCellText = CStr(cellEntry)
        End If
    Else
        TableCellText = CStr(cellEntry)
    End If
End Function

Private Function ErrorCellText(ByVal cellEntry As Variant) As String
    If cellEntry = CVErr(xlErrNA) Then
        ErrorCellText = "#N/A"
    ElseIf cellEntry = CVErr(xlErrDiv0) Then
        ErrorCellText = "#DIV/0!"
    ElseIf cellEntry = CVErr(xlErrValue) Then
        ErrorCellText = "#VALUE!"
    ElseIf cellEntry = CVErr(xlErrRef) Then
        ErrorCellText = "#REF!"
    ElseIf cellEntry = CVErr(xlErrName) Then
        ErrorCellText = "#NAME?"
    ElseIf cellEntry = CVErr(xlErrNum) Then
        ErrorCellText = "#NUM!"
    ElseIf cellEntry = CVErr(xlErrNull) Then
        ErrorCellText = "#NULL!"
    Else
        ErrorCellText = "#ERROR"
    End If
End Function

' Column edges: asked widths first, then the width left over shared by
' the other columns in proportion to their longest text among the
' header and the first 200 rows. A column holds numbers when every
' filled cell is a number or a date; its cells and header align right.
Private Sub LayoutTableColumns(ByVal textSize As Double)
    Dim layoutKey As String
    Dim colNo As Long
    Dim rowNo As Long
    Dim weights() As Double
    Dim weightSum As Double
    Dim fixedSum As Double
    Dim spareWidth As Double
    Dim edgeAt As Double
    Dim longest As Long
    Dim cellEntry As Variant

    layoutKey = CStr(mTableVersion) & ":" & CStr(mWidth) & ":" & CStr(textSize)
    For colNo = 1 To mTableCols
        layoutKey = layoutKey & ":" & CStr(mColumnPoints(colNo)) & "/" & mColumnFormats(colNo)
    Next colNo
    If layoutKey = mColumnLayoutKey Then Exit Sub
    mColumnLayoutKey = layoutKey
    If mTableCols = 0 Then Exit Sub
    ReDim mColumnEdges(1 To mTableCols)
    ReDim mColumnNumeric(1 To mTableCols)
    ReDim weights(1 To mTableCols)
    For colNo = 1 To mTableCols
        mColumnNumeric(colNo) = (mTableRows > 0)
        For rowNo = 1 To mTableRows
            cellEntry = mTableCells(colNo, rowNo)
            If Not IsBlankCell(cellEntry) And Not IsNumberCell(cellEntry) Then
                mColumnNumeric(colNo) = False
                Exit For
            End If
        Next rowNo
        If mColumnPoints(colNo) > 0 Then
            fixedSum = fixedSum + mColumnPoints(colNo)
        Else
            longest = Len(mTableHeads(colNo)) + 2
            For rowNo = 1 To mTableRows
                If rowNo > 200 Then Exit For
                If Len(TableCellText(colNo, mTableCells(colNo, rowNo))) > longest Then
                    longest = Len(TableCellText(colNo, mTableCells(colNo, rowNo)))
                End If
            Next rowNo
            If longest < 4 Then longest = 4
            If longest > 40 Then longest = 40
            weights(colNo) = longest
            weightSum = weightSum + longest
        End If
    Next colNo
    spareWidth = mWidth - 2 - fixedSum
    If spareWidth < 0 Then spareWidth = 0
    For colNo = 1 To mTableCols
        If mColumnPoints(colNo) > 0 Then
            edgeAt = edgeAt + mColumnPoints(colNo)
        ElseIf weightSum > 0 Then
            edgeAt = edgeAt + spareWidth * weights(colNo) / weightSum
        End If
        mColumnEdges(colNo) = edgeAt
    Next colNo
End Sub

Private Function ColumnLeftEdge(ByVal colNo As Long) As Double
    If colNo > 1 Then ColumnLeftEdge = mColumnEdges(colNo - 1)
End Function

' A cell's text cut to its column with an ellipsis when it would run
' into the next.
Private Function FitCellText( _
    ByVal cellWords As String, _
    ByVal cellSpan As Double, _
    ByVal textSize As Double _
) As String
    Dim fitChars As Long

    fitChars = Int((cellSpan - 2 * TABLE_CELL_PAD) / (textSize * 0.55))
    If Len(cellWords) > fitChars And fitChars >= 2 Then
        FitCellText = Left$(cellWords, fitChars - 1) & ChrW(8230)
    ElseIf Len(cellWords) > fitChars Then
        FitCellText = vbNullString
    Else
        FitCellText = cellWords
    End If
End Function

' A table: header cells over its columns, a sort arrow on the sorted
' one; rows windowed to its height, each one shape whose tab stops place
' the columns; and a footer that pages when the rows need more room.
' The selected row fills with the accent. Parts are rewritten only when
' their own look changed.
Private Sub ApplyTable(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim textSize As Double
    Dim rowPitch As Double
    Dim slotTotal As Long
    Dim colNo As Long
    Dim slotNo As Long
    Dim displayNo As Long
    Dim sourceRow As Long
    Dim headWords As String
    Dim rowWords As String
    Dim layoutKey As String
    Dim rowFill As Long
    Dim rowInk As Long
    Dim headFill As Long
    Dim headInk As Long

    textSize = mFontSize
    If textSize <= 0 Then textSize = themeValue.BaseFontSize
    rowPitch = textSize + 11
    EnsureTableShown
    LayoutTableColumns textSize
    slotTotal = Int((mHeight - 2 - rowPitch) / rowPitch)
    If slotTotal < 1 Then slotTotal = 1
    ' The footer shows while rows page or a filter holds.
    mTableFooterOn = (mTableShownCount > slotTotal) Or LenB(mTableFilter) > 0
    If mTableFooterOn Then
        slotTotal = Int((mHeight - 2 - 2 * rowPitch) / rowPitch)
        If slotTotal < 1 Then slotTotal = 1
    End If
    If mTableOffset > mTableShownCount - slotTotal Then
        mTableOffset = mTableShownCount - slotTotal
    End If
    If mTableOffset < 0 Then mTableOffset = 0
    mTableRowPitch = rowPitch
    mTableSlots = slotTotal
    headInk = ResolveDrawnInk(themeValue)
    If mTableCols > 0 Then EnsureRowStates mAppliedTableHeadStates, mTableCols
    For colNo = 1 To mTableCols
        headWords = mTableHeads(colNo)
        If colNo = mSortColumn Then
            headWords = headWords & " " & IIf(mSortDescending, ChrW(9660), ChrW(9650))
        End If
        headWords = FitCellText(headWords, mColumnEdges(colNo) - ColumnLeftEdge(colNo), _
            textSize)
        headFill = PointerTint(themeValue.MutedColor, themeValue.OnSurfaceColor, _
            "th" & colNo)
        SyncTableCell hostSheet, themeValue, mAppliedTableHeadStates(colNo), "th" & colNo, _
            mLeft + 1 + ColumnLeftEdge(colNo), mTop + 1, _
            mColumnEdges(colNo) - ColumnLeftEdge(colNo), rowPitch, headWords, _
            mColumnNumeric(colNo), headFill, headInk, True, textSize
    Next colNo
    For colNo = mTableCols + 1 To mAppliedTableColsDrawn
        DropPart hostSheet, "th" & colNo
    Next colNo
    mAppliedTableColsDrawn = mTableCols
    layoutKey = mColumnLayoutKey & ":" & CStr(mLeft) & ":" & CStr(mTop)
    EnsureRowStates mAppliedTableRowStates, slotTotal
    For slotNo = 1 To slotTotal
        displayNo = mTableOffset + slotNo
        If displayNo <= mTableShownCount Then
            sourceRow = mTableShown(displayNo)
            rowWords = TableRowWords(sourceRow, textSize)
            If sourceRow = CLng(mValue) Then
                rowFill = themeValue.PrimaryColor
                rowInk = themeValue.OnPrimaryColor
            Else
                rowFill = PointerTint(themeValue.SurfaceColor, themeValue.OnSurfaceColor, _
                    "tr" & slotNo)
                If rowFill = themeValue.SurfaceColor Then rowFill = -1
                rowInk = ResolveDrawnInk(themeValue)
            End If
        ElseIf displayNo = 1 Then
            If mTableRows > 0 Then
                rowWords = vbTab & Phrase(UI_NO_ROWS_MATCH)
            ElseIf LenB(mTableEmptyText) > 0 Then
                rowWords = vbTab & mTableEmptyText
            Else
                rowWords = vbTab & Phrase(UI_NO_ROWS)
            End If
            rowFill = -1
            rowInk = themeValue.OnMutedColor
        Else
            rowWords = vbNullString
        End If
        SyncTableRow hostSheet, themeValue, slotNo, layoutKey, rowWords, rowFill, _
            rowInk, rowPitch, textSize, (displayNo = 1 And mTableShownCount = 0)
    Next slotNo
    For slotNo = slotTotal + 1 To mAppliedTableSlotsDrawn
        DropPart hostSheet, "tr" & slotNo
    Next slotNo
    mAppliedTableSlotsDrawn = slotTotal
    SyncTableFooter hostSheet, themeValue, rowPitch, textSize
End Sub

' One row's cells joined for its shape: each cell after a tab, cut to
' its column, so the row's tab stops place them.
Private Function TableRowWords(ByVal sourceRow As Long, ByVal textSize As Double) As String
    Dim colNo As Long

    For colNo = 1 To mTableCols
        TableRowWords = TableRowWords & vbTab & FitCellText( _
            TableCellText(colNo, mTableCells(colNo, sourceRow)), _
            mColumnEdges(colNo) - ColumnLeftEdge(colNo), textSize)
    Next colNo
End Function

' A header cell: its own shape so a click knows the column.
Private Sub SyncTableCell( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByRef appliedKey As String, _
    ByVal partName As String, _
    ByVal partLeft As Double, _
    ByVal partTop As Double, _
    ByVal partWidth As Double, _
    ByVal partHeight As Double, _
    ByVal partWords As String, _
    ByVal alignRight As Boolean, _
    ByVal fillRgb As Long, _
    ByVal inkRgb As Long, _
    ByVal wordsBold As Boolean, _
    ByVal textSize As Double _
)
    Dim partKey As String
    Dim partShape As Shape

    partKey = PartsState(themeValue) & ":" & CStr(partLeft) & ":" & CStr(partWidth) & _
        ":" & CStr(partHeight) & ":" & partWords & ":" & CStr(alignRight) & ":" & _
        CStr(fillRgb) & ":" & CStr(inkRgb) & ":" & CStr(textSize)
    If appliedKey = partKey Then Exit Sub
    appliedKey = partKey
    Set partShape = EnsurePartShape(hostSheet, partName, msoShapeRectangle)
    partShape.Left = partLeft
    partShape.Top = partTop
    partShape.Width = partWidth
    partShape.Height = partHeight
    partShape.Fill.Visible = msoTrue
    partShape.Fill.ForeColor.RGB = fillRgb
    partShape.Line.Visible = msoFalse
    With partShape.TextFrame2
        .MarginLeft = TABLE_CELL_PAD
        .MarginRight = TABLE_CELL_PAD
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = partWords
        If alignRight Then
            .TextRange.ParagraphFormat.Alignment = msoAlignRight
        Else
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End If
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize
            .Bold = BoolToMso(wordsBold)
            .Fill.ForeColor.RGB = inkRgb
        End With
    End With
    partShape.Visible = BoolToMso(mVisible)
    partShape.OnAction = mAppliedOnAction
End Sub

' A row slot: one shape across the table with a thin rule; its tab stops
' change only when the columns do. A slot whose place, size, and font held
' rewrites only the words, fill, and ink that changed. An empty words
' string removes it.
Private Sub SyncTableRow( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal slotNo As Long, _
    ByVal layoutKey As String, _
    ByVal rowWords As String, _
    ByVal fillRgb As Long, _
    ByVal inkRgb As Long, _
    ByVal rowPitch As Double, _
    ByVal textSize As Double, _
    ByVal emptyNote As Boolean _
)
    Dim frameKey As String
    Dim lookKey As String
    Dim priorParts() As String
    Dim priorLook() As String
    Dim rowShape As Shape
    Dim colNo As Long

    If LenB(rowWords) = 0 Then
        If LenB(mAppliedTableRowStates(slotNo)) > 0 Then
            mAppliedTableRowStates(slotNo) = vbNullString
            DropPart hostSheet, "tr" & slotNo
        End If
        Exit Sub
    End If
    frameKey = layoutKey & Chr$(1) & PartsState(themeValue) & ":" & CStr(rowPitch) & _
        ":" & CStr(textSize) & ":" & CStr(emptyNote)
    lookKey = rowWords & Chr$(2) & CStr(fillRgb) & Chr$(2) & CStr(inkRgb)
    If mAppliedTableRowStates(slotNo) = frameKey & Chr$(3) & lookKey Then Exit Sub
    priorParts = Split(mAppliedTableRowStates(slotNo) & Chr$(3), Chr$(3))
    mAppliedTableRowStates(slotNo) = frameKey & Chr$(3) & lookKey
    Set rowShape = EnsurePartShape(hostSheet, "tr" & slotNo, msoShapeRectangle)
    If priorParts(0) = frameKey Then
        priorLook = Split(priorParts(1) & Chr$(2) & Chr$(2), Chr$(2))
        If priorLook(1) <> CStr(fillRgb) Then
            rowShape.Fill.Visible = BoolToMso(fillRgb >= 0)
            If fillRgb >= 0 Then rowShape.Fill.ForeColor.RGB = fillRgb
        End If
        With rowShape.TextFrame2.TextRange
            If priorLook(0) <> rowWords Then .Text = rowWords
            If priorLook(2) <> CStr(inkRgb) Or priorLook(0) <> rowWords Then
                .Font.Fill.ForeColor.RGB = inkRgb
            End If
        End With
        Exit Sub
    End If
    rowShape.Left = mLeft + 1
    rowShape.Top = mTop + 1 + slotNo * rowPitch
    rowShape.Width = mWidth - 2
    rowShape.Height = rowPitch
    rowShape.Fill.Visible = BoolToMso(fillRgb >= 0)
    If fillRgb >= 0 Then rowShape.Fill.ForeColor.RGB = fillRgb
    rowShape.Line.Visible = msoTrue
    rowShape.Line.ForeColor.RGB = themeValue.BorderColor
    rowShape.Line.Weight = 0.5
    With rowShape.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = rowWords
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        If Split(priorParts(0) & Chr$(1), Chr$(1))(0) <> layoutKey Then
            With .TextRange.ParagraphFormat.TabStops
                Do While .Count > 0
                    .Item(1).Clear
                Loop
                If emptyNote Then
                    .Add msoTabStopLeft, TABLE_CELL_PAD
                Else
                    For colNo = 1 To mTableCols
                        If mColumnNumeric(colNo) Then
                            .Add msoTabStopRight, mColumnEdges(colNo) - TABLE_CELL_PAD
                        Else
                            .Add msoTabStopLeft, ColumnLeftEdge(colNo) + TABLE_CELL_PAD
                        End If
                    Next colNo
                End If
            End With
        End If
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize
            .Bold = msoFalse
            .Fill.ForeColor.RGB = inkRgb
        End With
    End With
    rowShape.Visible = BoolToMso(mVisible)
    rowShape.OnAction = mAppliedOnAction
End Sub

' The footer while the rows outgrow the table: which rows show of how
' many, and arrows that page, muted at the ends.
Private Sub SyncTableFooter( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal rowPitch As Double, _
    ByVal textSize As Double _
)
    Dim footKey As String
    Dim footTop As Double
    Dim lastShown As Long
    Dim labelShape As Shape
    Dim footWords As String

    If mTableFooterOn Then
        lastShown = mTableOffset + mTableSlots
        If lastShown > mTableShownCount Then lastShown = mTableShownCount
        footWords = FooterWords(lastShown)
        footKey = PartsState(themeValue) & ":" & CStr(mTableOffset) & ":" & _
            CStr(lastShown) & ":" & CStr(mTableShownCount) & ":" & PointerKey() & _
            Chr$(1) & footWords
    End If
    If footKey = mAppliedTableFootSig Then Exit Sub
    mAppliedTableFootSig = footKey
    If LenB(footKey) = 0 Then
        DropPart hostSheet, "tf"
        DropPart hostSheet, "tp"
        DropPart hostSheet, "tn"
        Exit Sub
    End If
    footTop = mTop + mHeight - 1 - rowPitch
    Set labelShape = EnsurePartShape(hostSheet, "tf", msoShapeRectangle)
    labelShape.Left = mLeft + 1
    labelShape.Top = footTop
    labelShape.Width = mWidth - 2 - 2 * PAGER_TARGET - 8
    labelShape.Height = rowPitch
    labelShape.Fill.Visible = msoFalse
    labelShape.Line.Visible = msoFalse
    With labelShape.TextFrame2
        .MarginLeft = TABLE_CELL_PAD
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = footWords
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize - 1
            .Bold = msoFalse
            .Fill.ForeColor.RGB = themeValue.OnMutedColor
        End With
    End With
    labelShape.Visible = BoolToMso(mVisible)
    labelShape.OnAction = mAppliedOnAction
    SyncTablePager hostSheet, themeValue, "tp", ChrW(9664), _
        mLeft + mWidth - 2 * PAGER_TARGET - 6, footTop + (rowPitch - PAGER_TARGET) / 2, _
        mTableOffset > 0, textSize
    SyncTablePager hostSheet, themeValue, "tn", ChrW(9654), _
        mLeft + mWidth - PAGER_TARGET - 4, footTop + (rowPitch - PAGER_TARGET) / 2, _
        lastShown < mTableShownCount, textSize
End Sub

' "1-6 of 14", and with a filter "Filter "ap": 1-3 of 3"; no rows
' shown reads "0 of 0".
Private Function FooterWords(ByVal lastShown As Long) As String
    If mTableShownCount = 0 Then
        FooterWords = Phrase(UI_ROW_RANGE_EMPTY)
    Else
        FooterWords = Phrase(UI_ROW_RANGE, mTableOffset + 1, lastShown, mTableShownCount)
    End If
    If LenB(mTableFilter) > 0 Then
        FooterWords = Phrase(UI_FILTER_NOTE, mTableFilter, FooterWords)
    End If
End Function

Private Sub SyncTablePager( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal partName As String, _
    ByVal glyph As String, _
    ByVal partLeft As Double, _
    ByVal partTop As Double, _
    ByVal pagerOn As Boolean, _
    ByVal textSize As Double _
)
    Dim pagerShape As Shape
    Dim pagerFill As Long

    Set pagerShape = EnsurePartShape(hostSheet, partName, msoShapeRoundedRectangle)
    pagerShape.Left = partLeft
    pagerShape.Top = partTop
    pagerShape.Width = PAGER_TARGET
    pagerShape.Height = PAGER_TARGET
    pagerFill = themeValue.SurfaceColor
    If pagerOn Then
        pagerFill = PointerTint(themeValue.SurfaceColor, themeValue.OnSurfaceColor, partName)
    End If
    pagerShape.Fill.Visible = msoTrue
    pagerShape.Fill.ForeColor.RGB = pagerFill
    pagerShape.Line.Visible = msoFalse
    With pagerShape.TextFrame2
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
            .Size = textSize - 2
            If pagerOn And mEnabled Then
                .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
            Else
                .Fill.ForeColor.RGB = themeValue.BorderColor
            End If
        End With
    End With
    pagerShape.Visible = BoolToMso(mVisible)
    pagerShape.OnAction = mAppliedOnAction
End Sub

' The table part under a point, from the last layout: a header cell, a
' row slot, or a footer arrow.
Private Function TablePartAt(ByVal pointsX As Double, ByVal pointsY As Double) As String
    Dim colNo As Long
    Dim slotNo As Long

    If mTableRowPitch <= 0 Or mTableCols = 0 Then Exit Function
    If pointsY < mTop + 1 + mTableRowPitch Then
        For colNo = 1 To mTableCols
            If pointsX - mLeft - 1 < mColumnEdges(colNo) Then
                TablePartAt = "th" & colNo
                Exit Function
            End If
        Next colNo
        Exit Function
    End If
    If mTableFooterOn And pointsY >= mTop + mHeight - 1 - mTableRowPitch Then
        If pointsX >= mLeft + mWidth - PAGER_TARGET - 4 Then
            TablePartAt = "tn"
        ElseIf pointsX >= mLeft + mWidth - 2 * PAGER_TARGET - 6 Then
            TablePartAt = "tp"
        End If
        Exit Function
    End If
    slotNo = Int((pointsY - mTop - 1) / mTableRowPitch)
    If slotNo >= 1 And slotNo <= mTableSlots _
        And mTableOffset + slotNo <= mTableShownCount Then
        TablePartAt = "tr" & slotNo
    End If
End Function

' A header click sorts by its column, ascending first and then flipping;
' a row click selects its row; the footer arrows page. True when the
' selection changed, so the handlers run.
Private Function ClickTable(ByVal hostApp As ReDimUI, ByVal partName As String) As Boolean
    Dim partNumber As Long
    Dim sourceRow As Long

    If Len(partName) > 2 And IsNumeric(Mid$(partName, 3)) Then
        partNumber = CLng(Mid$(partName, 3))
    End If
    Select Case Left$(partName, 2)
        Case "th"
            If partNumber < 1 Or partNumber > mTableCols Then Exit Function
            If mSortColumn = partNumber Then
                mSortDescending = Not mSortDescending
            Else
                mSortColumn = partNumber
                mSortDescending = False
            End If
            mTableOrderCurrent = False
            mTableOffset = 0
            Touch
        Case "tr"
            EnsureTableShown
            If partNumber < 1 Or mTableOffset + partNumber > mTableShownCount Then Exit Function
            sourceRow = mTableShown(mTableOffset + partNumber)
            ClickTable = SelectTableRow(hostApp, sourceRow)
            ' The second click of a double click on the row opens it.
            If RowDoubleClicked("row" & sourceRow) Then RunRowOpen hostApp
        Case Else
            If partName = "tp" Then
                PageTable -1
            ElseIf partName = "tn" Then
                PageTable 1
            End If
    End Select
End Function

Private Sub PageTable(ByVal pageDirection As Long)
    Dim offsetBefore As Long

    offsetBefore = mTableOffset
    mTableOffset = mTableOffset + pageDirection * mTableSlots
    If mTableOffset > mTableShownCount - mTableSlots Then
        mTableOffset = mTableShownCount - mTableSlots
    End If
    If mTableOffset < 0 Then mTableOffset = 0
    If mTableOffset <> offsetBefore Then Touch
End Sub

' Selects a row by its source number: the first cell writes to WritesTo.
' Re-selecting the selected row changes nothing.
Private Function SelectTableRow(ByVal hostApp As ReDimUI, ByVal sourceRow As Long) As Boolean
    If sourceRow < 1 Or sourceRow > mTableRows Then Exit Function
    If sourceRow = CLng(mValue) Then Exit Function
    mValue = sourceRow
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mTableCells(1, sourceRow)
    Touch
    SelectTableRow = True
End Function

' Up and Down move the selection a row among the rows shown, Page Up and
' Page Down a page, Home and End to the first and last; the rows scroll
' to keep it in view, and each move fires OnChange. Typing filters the
' rows, Ctrl+C copies, and Enter opens the selected row when the table
' has OnRowOpen.
Private Sub HandleTableKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim shownAt As Long
    Dim aimAt As Long

    Select Case keyText
        Case "{ENTER}"
            If LenB(mRowOpenProc) > 0 And TableShownAt(CLng(mValue)) > 0 Then
                RunRowOpen hostApp
            Else
                hostApp.ClickDefaultButton
            End If
            Exit Sub
        Case "{COPY}"
            CopyTableRows
            Exit Sub
    End Select
    If TypeTableFilter(keyText) Then Exit Sub
    EnsureTableShown
    If mTableShownCount = 0 Then Exit Sub
    shownAt = TableShownAt(CLng(mValue))
    Select Case keyText
        Case "{UP}"
            aimAt = shownAt - 1
            If shownAt = 0 Then aimAt = 1
        Case "{DOWN}"
            aimAt = shownAt + 1
        Case "{PGUP}"
            aimAt = shownAt - mTableSlots
        Case "{PGDN}"
            aimAt = shownAt + mTableSlots
        Case "{HOME}"
            aimAt = 1
        Case "{END}"
            aimAt = mTableShownCount
        Case Else
            Exit Sub
    End Select
    If aimAt < 1 Then aimAt = 1
    If aimAt > mTableShownCount Then aimAt = mTableShownCount
    If aimAt <= mTableOffset Then mTableOffset = aimAt - 1
    If aimAt > mTableOffset + mTableSlots Then mTableOffset = aimAt - mTableSlots
    If SelectTableRow(hostApp, mTableShown(aimAt)) Then
        RunClickHandlers hostApp
    Else
        Touch
    End If
End Sub

Private Sub RunRowOpen(ByVal hostApp As ReDimUI)
    If LenB(mRowOpenProc) = 0 Then Exit Sub
    ReDimUI.RunHandlerAs hostApp, Me, mRowOpenProc
End Sub

Friend Sub JoinStack(ByVal memberId As String)
    If mStackMembers Is Nothing Then Set mStackMembers = New Collection
    mStackMembers.Add memberId
    MarkStackStale
End Sub

' Drops a member from the order; "" drops only members that no longer
' exist.
Friend Sub LeaveStack(ByVal memberId As String)
    Dim idx As Long
    Dim hostApp As ReDimUI
    Dim memberValue As ReDimUI

    If mStackMembers Is Nothing Then Exit Sub
    Set hostApp = OwnerApp
    For idx = mStackMembers.Count To 1 Step -1
        If LenB(memberId) = 0 Then
            If Not hostApp.TryGetComponent(CStr(mStackMembers.Item(idx)), memberValue) Then
                mStackMembers.Remove idx
            End If
        ElseIf StrComp(CStr(mStackMembers.Item(idx)), memberId, vbTextCompare) = 0 Then
            mStackMembers.Remove idx
        End If
    Next idx
    MarkStackStale
End Sub

Friend Sub MarkStackStale()
    mStackGen = mStackGen + 1
End Sub

' The room a control takes in its stack's flow: nothing while hidden;
' otherwise its size, with the caption row above it and the note row
' under a field that keeps one. A stack measures its members.
Friend Sub MeasureFlow( _
    ByVal themeValue As ReDimUI, _
    ByRef flowShown As Boolean, _
    ByRef flowWidth As Double, _
    ByRef flowHeight As Double, _
    ByRef rowAbove As Double, _
    ByRef rowBelow As Double _
)
    flowShown = mVisible
    If Not flowShown Then Exit Sub
    If mKind = KIND_STACK Then
        LayOutStack themeValue, 0, 0, False, flowWidth, flowHeight
    ElseIf mKind = KIND_BADGE Then
        flowWidth = BadgeSpan()
        flowHeight = mHeight
    Else
        flowWidth = mWidth
        flowHeight = mHeight
    End If
    rowAbove = CaptionRoom(themeValue)
    rowBelow = NoteRoom(themeValue)
End Sub

' The caption's row: a point smaller than the control's text, and its gap.
Private Function CaptionRoom(ByVal themeValue As ReDimUI) As Double
    If LenB(mCaptionText) = 0 Then Exit Function
    CaptionRoom = mFontSize
    If CaptionRoom <= 0 Then CaptionRoom = themeValue.BaseFontSize
    CaptionRoom = CaptionRoom + 7
End Function

' The note row under a field that keeps one: a hint, or a check, rule,
' or cap whose message or count shows there. Reserving it for the rule
' rather than the message keeps the stack still while a message comes
' and goes.
Private Function NoteRoom(ByVal themeValue As ReDimUI) As Double
    If LenB(mHintText) = 0 Then
        If Not IsFloatField() Then Exit Function
        If LenB(mValidateProc) = 0 And Not mRequired And LenB(mErrorText) = 0 _
            And mMaxLength = 0 Then Exit Function
    End If
    NoteRoom = mFontSize
    If NoteRoom <= 0 Then NoteRoom = themeValue.BaseFontSize
    NoteRoom = NoteRoom + 5
End Function

' Lays the members out from an origin: down, each shown member at the
' padding's left edge below the one before, or across, their tops below
' the tallest caption. With fillSlots the slots are kept for the
' members; the content's extent comes back either way, measured from
' the members' sizes alone, so a stack inside another measures without
' a place.
Private Sub LayOutStack( _
    ByVal themeValue As ReDimUI, _
    ByVal originLeft As Double, _
    ByVal originTop As Double, _
    ByVal fillSlots As Boolean, _
    ByRef contentWidth As Double, _
    ByRef contentHeight As Double _
)
    Dim hostApp As ReDimUI
    Dim memberValue As ReDimUI
    Dim memberId As Variant
    Dim flowShown As Boolean
    Dim flowWidth As Double
    Dim flowHeight As Double
    Dim rowAbove As Double
    Dim rowBelow As Double
    Dim maxAbove As Double
    Dim rowReach As Double
    Dim cursorAt As Double
    Dim innerWidth As Double
    Dim shownCount As Long

    contentWidth = 0
    contentHeight = 0
    If fillSlots Then mStackSlotCount = 0
    If mStackMembers Is Nothing Then
        contentWidth = 2 * mStackPadding
        contentHeight = 2 * mStackPadding
        Exit Sub
    End If
    Set hostApp = OwnerApp
    If fillSlots Then
        ReDim mStackSlotIds(1 To mStackMembers.Count + 1)
        ReDim mStackSlotLefts(1 To mStackMembers.Count + 1)
        ReDim mStackSlotTops(1 To mStackMembers.Count + 1)
        ReDim mStackSlotSpans(1 To mStackMembers.Count + 1)
    End If
    If mStackAcross Then
        For Each memberId In mStackMembers
            If hostApp.TryGetComponent(CStr(memberId), memberValue) Then
                memberValue.MeasureFlow themeValue, flowShown, flowWidth, flowHeight, _
                    rowAbove, rowBelow
                If flowShown And rowAbove > maxAbove Then maxAbove = rowAbove
            End If
        Next memberId
        cursorAt = originLeft + mStackPadding
    Else
        cursorAt = originTop + mStackPadding
        innerWidth = mWidth - 2 * mStackPadding
    End If
    For Each memberId In mStackMembers
        If hostApp.TryGetComponent(CStr(memberId), memberValue) Then
            memberValue.MeasureFlow themeValue, flowShown, flowWidth, flowHeight, _
                rowAbove, rowBelow
            If flowShown Then
                shownCount = shownCount + 1
                If mStackAcross Then
                    If fillSlots Then
                        mStackSlotIds(shownCount) = CStr(memberId)
                        mStackSlotLefts(shownCount) = cursorAt
                        mStackSlotTops(shownCount) = originTop + mStackPadding + maxAbove
                    End If
                    cursorAt = cursorAt + flowWidth + mStackGap
                    If flowHeight + rowBelow > rowReach Then rowReach = flowHeight + rowBelow
                Else
                    ' Stretch widens controls; a stack inside keeps its own
                    ' width, which its content sets across.
                    If mStackStretch And innerWidth > 0 And Not mStackAutoWidth _
                        And memberValue.ComponentKind <> KIND_STACK Then
                        flowWidth = innerWidth
                        If fillSlots Then mStackSlotSpans(shownCount) = innerWidth
                    End If
                    If fillSlots Then
                        mStackSlotIds(shownCount) = CStr(memberId)
                        mStackSlotLefts(shownCount) = originLeft + mStackPadding
                        mStackSlotTops(shownCount) = cursorAt + rowAbove
                    End If
                    cursorAt = cursorAt + rowAbove + flowHeight + rowBelow + mStackGap
                    If flowWidth > rowReach Then rowReach = flowWidth
                End If
            End If
        End If
    Next memberId
    If fillSlots Then mStackSlotCount = shownCount
    If mStackAcross Then
        contentWidth = cursorAt - originLeft + mStackPadding
        If shownCount > 0 Then contentWidth = contentWidth - mStackGap
        contentHeight = maxAbove + rowReach + 2 * mStackPadding
    Else
        contentHeight = cursorAt - originTop + mStackPadding
        If shownCount > 0 Then contentHeight = contentHeight - mStackGap
        contentWidth = rowReach + 2 * mStackPadding
        If Not mStackAutoWidth And mWidth > 0 Then contentWidth = mWidth
    End If
End Sub

' Lays the stack out at its place when a member, the stack, or its
' place changed since the last layout, and sizes the stack to the
' content: its height, and its width across or when given 0.
Private Sub ArrangeStack()
    Dim contentWidth As Double
    Dim contentHeight As Double

    If mStackLaidGen = mStackGen And mStackLaidLeft = mLeft And mStackLaidTop = mTop _
        And mStackLaidWidth = mWidth Then Exit Sub
    LayOutStack OwnerApp.Theme, mLeft, mTop, True, contentWidth, contentHeight
    If contentHeight < 1 Then contentHeight = 1
    mHeight = contentHeight
    If mStackAcross Or mStackAutoWidth Then
        If contentWidth < 1 Then contentWidth = 1
        mWidth = contentWidth
    End If
    mStackLaidGen = mStackGen
    mStackLaidLeft = mLeft
    mStackLaidTop = mTop
    mStackLaidWidth = mWidth
End Sub

' A member's slot from the last layout, with the width a stretching
' stack gives it (0 for its own); False while the member is hidden.
Friend Function StackSlot( _
    ByVal memberId As String, _
    ByRef slotLeft As Double, _
    ByRef slotTop As Double, _
    ByRef slotSpan As Double _
) As Boolean
    Dim slotNo As Long

    For slotNo = 1 To mStackSlotCount
        If StrComp(mStackSlotIds(slotNo), memberId, vbTextCompare) = 0 Then
            slotLeft = mStackSlotLefts(slotNo)
            slotTop = mStackSlotTops(slotNo)
            slotSpan = mStackSlotSpans(slotNo)
            StackSlot = True
            Exit Function
        End If
    Next slotNo
End Function

' Whether a member stands somewhere other than the slot given.
Friend Function SlotMoved( _
    ByVal slotLeft As Double, _
    ByVal slotTop As Double, _
    ByVal slotSpan As Double _
) As Boolean
    SlotMoved = Abs(mLeft - slotLeft) > 0.01 Or Abs(mTop - slotTop) > 0.01
    If slotSpan > 0 Then SlotMoved = SlotMoved Or Abs(mWidth - slotSpan) > 0.01
End Function

' After a stack draws: members whose slot changed apply at it, and when
' the stack showed or hid, every member applies to follow it.
Private Sub MoveStackMembers( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal hostApp As ReDimUI _
)
    Dim memberValue As ReDimUI
    Dim memberId As Variant
    Dim slotNo As Long

    If mStackMembers Is Nothing Then Exit Sub
    If mVisible <> mAppliedStackShown Then
        mAppliedStackShown = mVisible
        For Each memberId In mStackMembers
            If hostApp.TryGetComponent(CStr(memberId), memberValue) Then
                memberValue.Apply hostSheet, themeValue, hostApp
            End If
        Next memberId
        Exit Sub
    End If
    For slotNo = 1 To mStackSlotCount
        If hostApp.TryGetComponent(mStackSlotIds(slotNo), memberValue) Then
            If memberValue.SlotMoved(mStackSlotLefts(slotNo), mStackSlotTops(slotNo), _
                mStackSlotSpans(slotNo)) Then
                memberValue.Apply hostSheet, themeValue, hostApp
            End If
        End If
    Next slotNo
End Sub

' A queued reflow: the stack lays out again and draws, which moves the
' members whose place changed.
Friend Sub ReflowMembers( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal hostApp As ReDimUI _
)
    MarkStackStale
    Apply hostSheet, themeValue, hostApp
End Sub

' After a member draws: when the room it takes changed, its stack
' queues a reflow for the end of the flush.
Private Sub NoteFlow(ByVal hostApp As ReDimUI, ByVal themeValue As ReDimUI)
    Dim flowShown As Boolean
    Dim flowWidth As Double
    Dim flowHeight As Double
    Dim rowAbove As Double
    Dim rowBelow As Double
    Dim flowKey As String

    MeasureFlow themeValue, flowShown, flowWidth, flowHeight, rowAbove, rowBelow
    flowKey = CStr(flowShown) & ":" & CStr(flowWidth) & ":" & CStr(flowHeight) & _
        ":" & CStr(rowAbove) & ":" & CStr(rowBelow)
    If flowKey = mAppliedFlowSig Then Exit Sub
    mAppliedFlowSig = flowKey
    hostApp.QueueReflow mStackId
End Sub

' Opens or closes an expander from a click or a key: its panel follows,
' and WritesTo takes the new state.
Private Sub ToggleExpander(ByVal hostApp As ReDimUI)
    mChecked = Not mChecked
    If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
    Touch
End Sub

' A date without its time of day.
Private Function DayOnly(ByVal anyDate As Date) As Date
    DayOnly = DateSerial(Year(anyDate), Month(anyDate), Day(anyDate))
End Function

Private Function DayInRange(ByVal candidateDay As Date) As Boolean
    If mDateEarliest <> 0 And candidateDay < mDateEarliest Then Exit Function
    If mDateLatest <> 0 And candidateDay > mDateLatest Then Exit Function
    DayInRange = True
End Function

' Where the calendar card sits: under the face, or over it when the list
' rule found no room below.
Private Function CalendarTop() As Double
    If mListFlipped Then
        CalendarTop = mTop - CAL_DEPTH - 2
    Else
        CalendarTop = mTop + mHeight + 2
    End If
End Function

' The first day the grid shows: the start of the week that holds the
' month's first day, by the system's first day of the week.
Private Function CalendarGridStart() As Date
    CalendarGridStart = mCalendarMonth - (Weekday(mCalendarMonth, vbUseSystemDayOfWeek) - 1)
End Function

Private Sub OpenCalendar()
    Dim anchorDay As Date

    mSelectOpen = True
    If mValue <> 0 Then
        anchorDay = CDate(mValue)
    Else
        anchorDay = Date
    End If
    mCalendarMonth = DateSerial(Year(anchorDay), Month(anchorDay), 1)
    mCalendarCursor = anchorDay
    NoteListOpened
End Sub

' A click on the face or its caret opens or closes the calendar, an
' arrow turns the month, and a day picks it. True when the date changed,
' so the handlers run.
Private Function ClickCalendar(ByVal hostApp As ReDimUI, ByVal partName As String) As Boolean
    ' The week rows, the day marks, and the card lie over the days; the
    ' pointer says which day a click on them meant.
    If IsCalendarCover(partName) Then
        partName = CalendarPartUnderPointer()
        If LenB(partName) = 0 Then Exit Function
    End If
    Select Case partName
        Case vbNullString, "caret"
            If mSelectOpen Then
                mSelectOpen = False
            Else
                OpenCalendar
            End If
            Touch
        Case "cp", "cn"
            If Not mSelectOpen Then Exit Function
            mCalendarMonth = DateAdd("m", IIf(partName = "cp", -1, 1), mCalendarMonth)
            mCalendarCursor = DateSerial(Year(mCalendarMonth), Month(mCalendarMonth), _
                Day(mCalendarCursor))
            If Month(mCalendarCursor) <> Month(mCalendarMonth) Then
                mCalendarCursor = DateSerial(Year(mCalendarMonth), Month(mCalendarMonth) + 1, 0)
            End If
            Touch
        Case Else
            If Left$(partName, 2) <> "cd" Or Not IsNumeric(Mid$(partName, 3)) Then
                Exit Function
            End If
            If Not mSelectOpen Then Exit Function
            ClickCalendar = PickCalendarDay(hostApp, _
                CalendarGridStart() + CLng(Mid$(partName, 3)) - 1)
    End Select
End Function

' Picks a calendar day: the calendar closes, and a new date writes to
' WritesTo and returns True for OnChange. A day outside DateRange is not
' picked.
Private Function PickCalendarDay(ByVal hostApp As ReDimUI, ByVal pickedDay As Date) As Boolean
    If Not DayInRange(pickedDay) Then Exit Function
    mSelectOpen = False
    If CDbl(pickedDay) <> mValue Then
        mValue = CDbl(pickedDay)
        If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, pickedDay
        PickCalendarDay = True
    End If
    Touch
End Function

' Closed, Alt+Down, F4, Space, and Down open the calendar on the date
' held or today. Open, the arrows move a day or a week, Page Up and
' Page Down a month, Home and End to the month's first and last day,
' and Enter or Space picks the day the keys reached.
Private Sub HandleDateKey(ByVal hostApp As ReDimUI, ByVal keyText As String)
    Dim aimDay As Date

    If Not mSelectOpen Then
        Select Case keyText
            Case "{ALTDOWN}", "{F4}", " ", "{DOWN}"
                OpenCalendar
                Touch
            Case "{ENTER}"
                hostApp.ClickDefaultButton
        End Select
        Exit Sub
    End If
    aimDay = mCalendarCursor
    Select Case keyText
        Case "{LEFT}"
            aimDay = aimDay - 1
        Case "{RIGHT}"
            aimDay = aimDay + 1
        Case "{UP}"
            aimDay = aimDay - 7
        Case "{DOWN}"
            aimDay = aimDay + 7
        Case "{PGUP}"
            aimDay = DateAdd("m", -1, aimDay)
        Case "{PGDN}"
            aimDay = DateAdd("m", 1, aimDay)
        Case "{HOME}"
            aimDay = DateSerial(Year(aimDay), Month(aimDay), 1)
        Case "{END}"
            aimDay = DateSerial(Year(aimDay), Month(aimDay) + 1, 0)
        Case "{ENTER}", " "
            If PickCalendarDay(hostApp, mCalendarCursor) Then RunClickHandlers hostApp
            Exit Sub
        Case "{F4}", "{ALTUP}"
            CloseDropList
            Exit Sub
        Case Else
            Exit Sub
    End Select
    mCalendarCursor = aimDay
    mCalendarMonth = DateSerial(Year(aimDay), Month(aimDay), 1)
    Touch
End Sub

' The calendar part under a point: the month arrows, or a day cell.
Private Function CalendarPartAt(ByVal pointsX As Double, ByVal pointsY As Double) As String
    Dim gridX As Double
    Dim gridY As Double

    gridX = pointsX - mLeft - CAL_PAD
    gridY = pointsY - CalendarTop() - CAL_PAD
    If gridX < 0 Or gridX >= 7 * CAL_CELL_W Or gridY < 0 Then Exit Function
    If gridY < CAL_HEAD_H Then
        If gridX < CAL_CELL_W Then
            CalendarPartAt = "cp"
        ElseIf gridX >= 6 * CAL_CELL_W Then
            CalendarPartAt = "cn"
        End If
        Exit Function
    End If
    gridY = gridY - CAL_HEAD_H - CAL_WEEK_H
    If gridY < 0 Or gridY >= 6 * CAL_CELL_H Then Exit Function
    CalendarPartAt = "cd" & (Int(gridY / CAL_CELL_H) * 7 + Int(gridX / CAL_CELL_W) + 1)
End Function

' True for a calendar part that lies over the days rather than naming
' one: the card, the weekday line, a week, or a day mark.
Private Function IsCalendarCover(ByVal partName As String) As Boolean
    Select Case partName
        Case "cb", "cw", "cs", "cv", "ct", "cc", "cr1", "cr2", "cr3", "cr4", "cr5", "cr6"
            IsCalendarCover = True
    End Select
End Function

' The calendar part under the pointer: an arrow, a day, or nothing.
Private Function CalendarPartUnderPointer() As String
    Dim pointsX As Double
    Dim pointsY As Double

    If Not mSelectOpen Then Exit Function
    If Not PointerPoints(pointsX, pointsY) Then Exit Function
    CalendarPartUnderPointer = CalendarPartAt(pointsX, pointsY)
End Function

' The open calendar: a card under the face (over it when there is no room
' below) with the month between its arrows, the weekday initials, and six
' weeks of days. A week is one line of text with a center tab stop over
' each column, where a shape per day cost the opening 36 more creations,
' and the second to sixth weeks start as copies of the first. Marks sit on
' single days: fills behind the weeks for the date held and the day under
' the pointer, rings over them for today and, while the picker has the
' keys, a dashed one for the day the keys reached. Days of other months or
' outside DateRange read muted, the held date's digits take the accent's
' ink, and today's are bold. Each part is rewritten only when its own look
' changed.
Private Sub SyncCalendar(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim cardTop As Double
    Dim gridTop As Double
    Dim textSize As Double
    Dim gridStart As Date
    Dim heldCell As Long
    Dim todayCell As Long
    Dim cursorCell As Long
    Dim hoverCell As Long
    Dim hoverFill As Long
    Dim weekNo As Long
    Dim partShape As Shape

    If Not (mSelectOpen And mVisible) Then
        ' A first draw looks once for a card an earlier session left open,
        ' as a VBA reset leaves it, and sweeps that calendar away.
        If Not mAppliedValid Then
            On Error Resume Next
            Set partShape = hostSheet.Shapes(ShapeName & PART_SEPARATOR & "cb")
            On Error GoTo 0
            If Not partShape Is Nothing Then
                EnsureRowStates mAppliedRowStates, CAL_SLOTS
                mAppliedRowStates(1) = "?"
            End If
        End If
        DeleteCalendarParts hostSheet
        Exit Sub
    End If
    EnsureRowStates mAppliedRowStates, CAL_SLOTS
    textSize = mFontSize
    If textSize <= 0 Then textSize = themeValue.BaseFontSize
    cardTop = CalendarTop()
    If mAppliedRowStates(1) <> PartsState(themeValue) & ":" & CStr(cardTop) Then
        mAppliedRowStates(1) = PartsState(themeValue) & ":" & CStr(cardTop)
        Set partShape = EnsurePartShape(hostSheet, "cb", msoShapeRoundedRectangle)
        On Error Resume Next
        partShape.Adjustments(1) = 0.04
        On Error GoTo 0
        partShape.Left = mLeft
        partShape.Top = cardTop
        partShape.Width = CAL_WIDTH
        partShape.Height = CAL_DEPTH
        partShape.Fill.Visible = msoTrue
        partShape.Fill.ForeColor.RGB = themeValue.SurfaceColor
        partShape.Line.Visible = msoTrue
        partShape.Line.ForeColor.RGB = themeValue.BorderColor
        partShape.Line.Weight = 1
        partShape.OnAction = mAppliedOnAction
    End If
    SyncCalendarPart hostSheet, themeValue, 2, "ch", mLeft + CAL_PAD + CAL_CELL_W, _
        cardTop + CAL_PAD, 5 * CAL_CELL_W, CAL_HEAD_H, _
        Format$(mCalendarMonth, "mmmm yyyy"), -1, themeValue.OnSurfaceColor, -1, _
        True, textSize
    SyncCalendarPart hostSheet, themeValue, 3, "cp", mLeft + CAL_PAD, _
        cardTop + CAL_PAD, CAL_CELL_W, CAL_HEAD_H, ChrW(9664), _
        PartFill("cp", themeValue), themeValue.OnMutedColor, -1, False, textSize - 2
    SyncCalendarPart hostSheet, themeValue, 4, "cn", mLeft + CAL_PAD + 6 * CAL_CELL_W, _
        cardTop + CAL_PAD, CAL_CELL_W, CAL_HEAD_H, ChrW(9654), _
        PartFill("cn", themeValue), themeValue.OnMutedColor, -1, False, textSize - 2
    gridStart = CalendarGridStart()
    gridTop = cardTop + CAL_PAD + CAL_HEAD_H + CAL_WEEK_H
    If mValue <> 0 Then heldCell = CalendarCellOf(CDate(mValue), gridStart)
    todayCell = CalendarCellOf(Date, gridStart)
    If ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) Then
        cursorCell = CalendarCellOf(mCalendarCursor, gridStart)
    End If
    If Left$(mHoveredPart, 2) = "cd" Then
        hoverFill = PartFill(mHoveredPart, themeValue)
        If hoverFill >= 0 Then hoverCell = CLng(Val(Mid$(mHoveredPart, 3)))
        If hoverCell = heldCell Then hoverCell = 0
    End If
    ' The fills go down before the weeks, so the digits draw over them.
    SyncCalendarMark hostSheet, themeValue, CAL_SLOT_HELD, "cs", heldCell, gridTop, _
        themeValue.PrimaryColor, -1, False, True
    SyncCalendarMark hostSheet, themeValue, CAL_SLOT_HOVER, "cv", hoverCell, gridTop, _
        hoverFill, -1, False, True
    SyncCalendarWeekdays hostSheet, themeValue, cardTop, gridStart, textSize
    For weekNo = 1 To 6
        SyncCalendarWeek hostSheet, themeValue, weekNo, gridStart, gridTop, textSize - 1, _
            heldCell, todayCell
    Next weekNo
    ' The rings go over the weeks, clear of the digits. The cursor's ring
    ' stands in for today's, and the held date takes no today ring.
    If todayCell = heldCell Or todayCell = cursorCell Then todayCell = 0
    SyncCalendarMark hostSheet, themeValue, CAL_SLOT_TODAY, "ct", todayCell, gridTop, _
        -1, themeValue.PrimaryColor, False, False
    SyncCalendarMark hostSheet, themeValue, CAL_SLOT_CURSOR, "cc", cursorCell, gridTop, _
        -1, themeValue.OnSurfaceColor, True, False
End Sub

' The grid cell, 1 to 42, that shows a day; 0 when the six weeks miss it.
Private Function CalendarCellOf(ByVal shownDay As Date, ByVal gridStart As Date) As Long
    Dim dayIndex As Long

    dayIndex = CLng(DayOnly(shownDay) - gridStart)
    If dayIndex >= 0 And dayIndex < 42 Then CalendarCellOf = dayIndex + 1
End Function

' A mark on one day cell: a fill behind the weeks or a ring over them,
' hidden for cell 0. A fill is drawn with the calendar even while hidden,
' so the weeks stack over it; a ring waits for a day that needs one.
Private Sub SyncCalendarMark( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal stateSlot As Long, _
    ByVal partName As String, _
    ByVal cellNo As Long, _
    ByVal gridTop As Double, _
    ByVal fillRgb As Long, _
    ByVal ringRgb As Long, _
    ByVal ringDashed As Boolean, _
    ByVal behindWeeks As Boolean _
)
    Dim lookKey As String
    Dim placeKey As String
    Dim priorParts() As String
    Dim markShape As Shape
    Dim markMade As Boolean

    If cellNo = 0 Then
        Select Case mAppliedRowStates(stateSlot)
            Case "-"
                Exit Sub
            Case vbNullString
                If Not behindWeeks Then Exit Sub
        End Select
        Set markShape = EnsureCalendarMark(hostSheet, partName, markMade)
        markShape.Visible = msoFalse
        mAppliedRowStates(stateSlot) = "-"
        Exit Sub
    End If
    lookKey = CStr(ObjPtr(themeValue)) & ":" & CStr(fillRgb) & ":" & CStr(ringRgb) & ":" & _
        CStr(ringDashed) & ":" & mAppliedOnAction
    placeKey = CStr(cellNo) & ":" & CStr(mLeft) & ":" & CStr(gridTop)
    If mAppliedRowStates(stateSlot) = lookKey & Chr$(3) & placeKey Then Exit Sub
    priorParts = Split(mAppliedRowStates(stateSlot) & Chr$(3), Chr$(3))
    Set markShape = EnsureCalendarMark(hostSheet, partName, markMade)
    If priorParts(0) = "-" Then markShape.Visible = msoTrue
    mAppliedRowStates(stateSlot) = lookKey & Chr$(3) & placeKey
    markShape.Left = mLeft + CAL_PAD + ((cellNo - 1) Mod 7) * CAL_CELL_W + 1
    markShape.Top = gridTop + ((cellNo - 1) \ 7) * CAL_CELL_H + 1
    If priorParts(0) = lookKey And Not markMade Then Exit Sub
    markShape.Fill.Visible = BoolToMso(fillRgb >= 0)
    If fillRgb >= 0 Then markShape.Fill.ForeColor.RGB = fillRgb
    ApplyCalendarRing markShape, ringRgb, ringDashed
    markShape.OnAction = mAppliedOnAction
End Sub

' A day mark's shape: found, or made at a day cell's size.
Private Function EnsureCalendarMark( _
    ByVal hostSheet As Worksheet, _
    ByVal partName As String, _
    ByRef markMade As Boolean _
) As Shape
    Dim markShape As Shape

    Set markShape = EnsurePartShape(hostSheet, partName, msoShapeRoundedRectangle, markMade)
    If markMade Then
        On Error Resume Next
        markShape.Adjustments(1) = 0.3
        On Error GoTo 0
        markShape.Width = CAL_CELL_W - 2
        markShape.Height = CAL_CELL_H - 2
    End If
    Set EnsureCalendarMark = markShape
End Function

' One week of days on one line: a tab before each day's number. A week
' rewrites its words and runs only when they changed: the plain ink
' first, then the muted runs, the held date's ink, and today's bold.
Private Sub SyncCalendarWeek( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal weekNo As Long, _
    ByVal gridStart As Date, _
    ByVal gridTop As Double, _
    ByVal textSize As Double, _
    ByVal heldCell As Long, _
    ByVal todayCell As Long _
)
    Dim weekWords As String
    Dim dayStarts(1 To 7) As Long
    Dim dayEnds(1 To 7) As Long
    Dim mutedDays As String
    Dim columnNo As Long
    Dim runFrom As Long
    Dim cellDay As Date
    Dim heldColumn As Long
    Dim todayColumn As Long
    Dim frameKey As String
    Dim lookKey As String
    Dim stateSlot As Long
    Dim weekShape As Shape
    Dim weekMade As Boolean

    For columnNo = 1 To 7
        cellDay = gridStart + (weekNo - 1) * 7 + columnNo - 1
        weekWords = weekWords & vbTab
        dayStarts(columnNo) = Len(weekWords) + 1
        weekWords = weekWords & CStr(Day(cellDay))
        dayEnds(columnNo) = Len(weekWords)
        If Month(cellDay) = Month(mCalendarMonth) And DayInRange(cellDay) Then
            mutedDays = mutedDays & "0"
        Else
            mutedDays = mutedDays & "1"
        End If
    Next columnNo
    If heldCell > 0 Then
        If (heldCell - 1) \ 7 + 1 = weekNo Then heldColumn = (heldCell - 1) Mod 7 + 1
    End If
    If todayCell > 0 Then
        If (todayCell - 1) \ 7 + 1 = weekNo Then todayColumn = (todayCell - 1) Mod 7 + 1
    End If
    stateSlot = CAL_SLOT_WEEK1 + weekNo - 1
    frameKey = CStr(ObjPtr(themeValue)) & ":" & CStr(mLeft) & ":" & CStr(gridTop) & ":" & _
        CStr(textSize) & ":" & mAppliedOnAction
    lookKey = weekWords & Chr$(2) & mutedDays & Chr$(2) & CStr(heldColumn) & Chr$(2) & _
        CStr(todayColumn) & Chr$(2) & CStr(themeValue.OnSurfaceColor) & ":" & _
        CStr(themeValue.OnMutedColor) & ":" & CStr(themeValue.OnPrimaryColor)
    If mAppliedRowStates(stateSlot) = frameKey & Chr$(3) & lookKey Then Exit Sub
    If Left$(mAppliedRowStates(stateSlot), Len(frameKey) + 1) = frameKey & Chr$(3) Then
        Set weekShape = EnsurePartShape(hostSheet, "cr" & weekNo, msoShapeRectangle, weekMade)
    End If
    If weekShape Is Nothing Or weekMade Then
        Set weekShape = PlaceCalendarWeek(hostSheet, themeValue, weekNo, gridTop, textSize, _
            frameKey)
    End If
    mAppliedRowStates(stateSlot) = frameKey & Chr$(3) & lookKey
    With weekShape.TextFrame2.TextRange
        .Text = weekWords
        .Font.Fill.ForeColor.RGB = themeValue.OnSurfaceColor
        .Font.Bold = msoFalse
        columnNo = 1
        Do While columnNo <= 7
            If Mid$(mutedDays, columnNo, 1) = "1" Then
                runFrom = columnNo
                Do While columnNo < 7
                    If Mid$(mutedDays, columnNo + 1, 1) <> "1" Then Exit Do
                    columnNo = columnNo + 1
                Loop
                .Characters(dayStarts(runFrom), dayEnds(columnNo) - dayStarts(runFrom) + 1) _
                    .Font.Fill.ForeColor.RGB = themeValue.OnMutedColor
            End If
            columnNo = columnNo + 1
        Loop
        If heldColumn > 0 Then
            .Characters(dayStarts(heldColumn), dayEnds(heldColumn) - dayStarts(heldColumn) + 1) _
                .Font.Fill.ForeColor.RGB = themeValue.OnPrimaryColor
        End If
        If todayColumn > 0 Then
            .Characters(dayStarts(todayColumn), _
                dayEnds(todayColumn) - dayStarts(todayColumn) + 1).Font.Bold = msoTrue
        End If
    End With
End Sub

' Places a week's line. A fresh week after the first copies the first,
' placed just before it in the same frame: the copy carries the margins,
' tab stops, font, and click, and costs a fraction of drawing them.
Private Function PlaceCalendarWeek( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal weekNo As Long, _
    ByVal gridTop As Double, _
    ByVal textSize As Double, _
    ByVal frameKey As String _
) As Shape
    Dim weekShape As Shape
    Dim columnNo As Long

    If weekNo > 1 And LenB(mAppliedRowStates(CAL_SLOT_WEEK1 + weekNo - 1)) = 0 Then
        If Left$(mAppliedRowStates(CAL_SLOT_WEEK1), Len(frameKey) + 1) = frameKey & Chr$(3) Then
            Set weekShape = CopyPartShape(hostSheet, _
                hostSheet.Shapes(ShapeName & PART_SEPARATOR & "cr1"), "cr" & weekNo)
            If Not weekShape Is Nothing Then
                weekShape.Left = mLeft + CAL_PAD
                weekShape.Top = gridTop + (weekNo - 1) * CAL_CELL_H + 1
                Set PlaceCalendarWeek = weekShape
                Exit Function
            End If
        End If
    End If
    Set weekShape = EnsurePartShape(hostSheet, "cr" & weekNo, msoShapeRectangle)
    weekShape.Left = mLeft + CAL_PAD
    weekShape.Top = gridTop + (weekNo - 1) * CAL_CELL_H + 1
    weekShape.Width = 7 * CAL_CELL_W
    weekShape.Height = CAL_CELL_H - 2
    weekShape.Fill.Visible = msoFalse
    With weekShape.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange.ParagraphFormat
            .Alignment = msoAlignLeft
            With .TabStops
                Do While .Count > 0
                    .Item(1).Clear
                Loop
                For columnNo = 1 To 7
                    .Add msoTabStopCenter, (columnNo - 0.5) * CAL_CELL_W
                Next columnNo
            End With
        End With
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize
        End With
    End With
    weekShape.OnAction = mAppliedOnAction
    Set PlaceCalendarWeek = weekShape
End Function

' The part a calendar state slot records, for the teardown.
Private Function CalendarSlotPart(ByVal stateSlot As Long) As String
    Select Case stateSlot
        Case CAL_SLOT_HELD
            CalendarSlotPart = "cs"
        Case CAL_SLOT_HOVER
            CalendarSlotPart = "cv"
        Case CAL_SLOT_TODAY
            CalendarSlotPart = "ct"
        Case CAL_SLOT_CURSOR
            CalendarSlotPart = "cc"
        Case Else
            CalendarSlotPart = "cr" & (stateSlot - CAL_SLOT_WEEK1 + 1)
    End Select
End Function

' The weekday initials: one shape whose center tab stops sit over the
' day columns, where seven shapes cost the opening six more creations.
Private Sub SyncCalendarWeekdays( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal cardTop As Double, _
    ByVal gridStart As Date, _
    ByVal textSize As Double _
)
    Dim weekWords As String
    Dim weekKey As String
    Dim dayNo As Long
    Dim weekShape As Shape

    For dayNo = 1 To 7
        weekWords = weekWords & vbTab & _
            Left$(WeekdayName(Weekday(gridStart + dayNo - 1), True), 2)
    Next dayNo
    weekKey = PartsState(themeValue) & ":" & CStr(cardTop) & ":" & weekWords & _
        ":" & CStr(textSize)
    If mAppliedRowStates(5) = weekKey Then Exit Sub
    mAppliedRowStates(5) = weekKey
    Set weekShape = EnsurePartShape(hostSheet, "cw", msoShapeRectangle)
    weekShape.Left = mLeft + CAL_PAD
    weekShape.Top = cardTop + CAL_PAD + CAL_HEAD_H
    weekShape.Width = 7 * CAL_CELL_W
    weekShape.Height = CAL_WEEK_H
    weekShape.Fill.Visible = msoFalse
    weekShape.Line.Visible = msoFalse
    With weekShape.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = weekWords
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.ParagraphFormat.TabStops
            Do While .Count > 0
                .Item(1).Clear
            Loop
            For dayNo = 1 To 7
                .Add msoTabStopCenter, (dayNo - 0.5) * CAL_CELL_W
            Next dayNo
        End With
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize - 2
            .Bold = msoFalse
            .Fill.ForeColor.RGB = themeValue.OnMutedColor
        End With
    End With
    weekShape.OnAction = mAppliedOnAction
End Sub

' The hover look of a calendar part: a tint of the card under PointerEffects,
' otherwise no fill (-1).
Private Function PartFill(ByVal partName As String, ByVal themeValue As ReDimUI) As Long
    PartFill = PointerTint(themeValue.SurfaceColor, themeValue.OnSurfaceColor, partName)
    If PartFill = themeValue.SurfaceColor Then PartFill = -1
End Function

' One calendar header part, the month or an arrow: a rounded cell with
' centered text, written only when its look differs from what its slot
' recorded. A part whose place, size, and font held rewrites only the
' words, fill, ink, and ring that changed, as the month does when it turns.
Private Sub SyncCalendarPart( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal stateSlot As Long, _
    ByVal partName As String, _
    ByVal partLeft As Double, _
    ByVal partTop As Double, _
    ByVal partWidth As Double, _
    ByVal partHeight As Double, _
    ByVal partWords As String, _
    ByVal fillRgb As Long, _
    ByVal inkRgb As Long, _
    ByVal ringRgb As Long, _
    ByVal wordsBold As Boolean, _
    ByVal textSize As Double, _
    Optional ByVal ringDashed As Boolean _
)
    Dim frameKey As String
    Dim lookKey As String
    Dim priorParts() As String
    Dim priorLook() As String
    Dim partShape As Shape

    frameKey = CStr(ObjPtr(themeValue)) & ":" & CStr(partLeft) & ":" & CStr(partTop) & _
        ":" & CStr(partWidth) & ":" & CStr(partHeight) & ":" & CStr(textSize) & ":" & _
        mAppliedOnAction
    lookKey = partWords & Chr$(2) & CStr(fillRgb) & Chr$(2) & CStr(inkRgb) & Chr$(2) & _
        CStr(ringRgb) & Chr$(2) & CStr(wordsBold) & Chr$(2) & CStr(ringDashed)
    If mAppliedRowStates(stateSlot) = frameKey & Chr$(3) & lookKey Then Exit Sub
    priorParts = Split(mAppliedRowStates(stateSlot) & Chr$(3), Chr$(3))
    mAppliedRowStates(stateSlot) = frameKey & Chr$(3) & lookKey
    Set partShape = EnsurePartShape(hostSheet, partName, msoShapeRoundedRectangle)
    If priorParts(0) = frameKey Then
        priorLook = Split(priorParts(1) & String$(5, Chr$(2)), Chr$(2))
        If priorLook(1) <> CStr(fillRgb) Then
            partShape.Fill.Visible = BoolToMso(fillRgb >= 0)
            If fillRgb >= 0 Then partShape.Fill.ForeColor.RGB = fillRgb
        End If
        If priorLook(3) <> CStr(ringRgb) Or priorLook(5) <> CStr(ringDashed) Then
            ApplyCalendarRing partShape, ringRgb, ringDashed
        End If
        With partShape.TextFrame2.TextRange
            If priorLook(0) <> partWords Then .Text = partWords
            If priorLook(2) <> CStr(inkRgb) Or priorLook(0) <> partWords Then
                .Font.Fill.ForeColor.RGB = inkRgb
            End If
            If priorLook(4) <> CStr(wordsBold) Or priorLook(0) <> partWords Then
                .Font.Bold = BoolToMso(wordsBold)
            End If
        End With
        Exit Sub
    End If
    On Error Resume Next
    partShape.Adjustments(1) = 0.3
    On Error GoTo 0
    partShape.Left = partLeft
    partShape.Top = partTop
    partShape.Width = partWidth
    partShape.Height = partHeight
    partShape.Fill.Visible = BoolToMso(fillRgb >= 0)
    If fillRgb >= 0 Then partShape.Fill.ForeColor.RGB = fillRgb
    ApplyCalendarRing partShape, ringRgb, ringDashed
    With partShape.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = partWords
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = textSize
            .Bold = BoolToMso(wordsBold)
            .Fill.ForeColor.RGB = inkRgb
        End With
    End With
    partShape.OnAction = mAppliedOnAction
End Sub

Private Sub ApplyCalendarRing( _
    ByVal partShape As Shape, _
    ByVal ringRgb As Long, _
    ByVal ringDashed As Boolean _
)
    partShape.Line.Visible = BoolToMso(ringRgb >= 0)
    If ringRgb < 0 Then Exit Sub
    partShape.Line.ForeColor.RGB = ringRgb
    partShape.Line.Weight = 1.25
    If ringDashed Then
        partShape.Line.DashStyle = msoLineDash
    Else
        partShape.Line.DashStyle = msoLineSolid
    End If
End Sub

' A closed calendar leaves no shapes behind. A close drops the parts its
' slots say were drawn; the sweep of a calendar an earlier session left
' open drops every name, and the day cells releases before 1.0.1 drew.
Private Sub DeleteCalendarParts(ByVal hostSheet As Worksheet)
    Dim partNo As Long
    Dim partHeld As Boolean
    Dim sweeping As Boolean

    On Error Resume Next
    partHeld = (LenB(mAppliedRowStates(1)) > 0)
    sweeping = (mAppliedRowStates(1) = "?")
    On Error GoTo 0
    If Not partHeld Then Exit Sub
    ' The picker stops ticking with its calendar closed.
    OwnerApp.TickablesChanged
    On Error Resume Next
    DropPart hostSheet, "cb"
    DropPart hostSheet, "ch"
    DropPart hostSheet, "cp"
    DropPart hostSheet, "cn"
    DropPart hostSheet, "cw"
    For partNo = CAL_SLOT_HELD To CAL_SLOTS
        If sweeping Or LenB(mAppliedRowStates(partNo)) > 0 Then
            DropPart hostSheet, CalendarSlotPart(partNo)
        End If
    Next partNo
    If sweeping Then
        For partNo = 1 To 42
            DropPart hostSheet, "cd" & partNo
        Next partNo
    End If
    On Error GoTo 0
    Erase mAppliedRowStates
End Sub

' The tab a Tabs control shows: its value held to its tabs, the first
' when the value is out of range, and 0 with no tabs.
Private Function ShownTab() As Long
    Dim tabTotal As Long

    tabTotal = ItemCount
    If tabTotal = 0 Then Exit Function
    ShownTab = CLng(mValue)
    If ShownTab < 1 Or ShownTab > tabTotal Then ShownTab = 1
End Function

' For the panels: the tab shown, 0 while the Tabs control is hidden, and
' -1 for any other kind.
Friend Property Get PanelTab() As Long
    If mKind = KIND_TABS Then
        If mVisible Then PanelTab = ShownTab()
    ElseIf mKind = KIND_EXPANDER Then
        ' An open expander shows its one panel.
        If mVisible And mChecked Then PanelTab = 1
    Else
        PanelTab = -1
    End If
End Property

Friend Property Get TabOwnerId() As String
    TabOwnerId = mTabOwnerId
End Property

Friend Property Get TabShut() As Boolean
    TabShut = mTabShut
End Property

' Whether this control's tab is hidden while panelNow shows; a panel of
' no Tabs control (-1) is never hidden.
Friend Function TabShutUnder(ByVal panelNow As Long) As Boolean
    If panelNow < 0 Then Exit Function
    TabShutUnder = (mTabNumber <> panelNow)
End Function

' Hides the control while its tab is not the one shown, keeping the
' visibility it has apart from the tab, and gives it back when its tab
' shows. A list open on a panel that hides closes.
Private Sub RefreshTabShut(ByVal hostApp As ReDimUI)
    Dim shutNow As Boolean

    If LenB(mTabOwnerId) = 0 And LenB(mStackId) = 0 And Not mTabShut Then Exit Sub
    If LenB(mTabOwnerId) > 0 Then
        shutNow = TabShutUnder(hostApp.ShownTabOf(mTabOwnerId))
    End If
    ' A hidden stack hides its members as a hidden tab does.
    If Not shutNow And LenB(mStackId) > 0 Then
        shutNow = hostApp.StackHidden(mStackId)
    End If
    If shutNow = mTabShut Then Exit Sub
    If shutNow Then
        mVisibleBeforeTab = mVisible
        mVisible = False
        mSelectOpen = False
        mListHighlight = 0
    Else
        mVisible = mVisibleBeforeTab
    End If
    mTabShut = shutNow
End Sub

' After a Tabs control draws: when the tab its panels show for changed,
' the controls that change with it re-apply.
Private Sub SyncTabPanels(ByVal hostApp As ReDimUI)
    Dim panelNow As Long

    panelNow = PanelTab
    If panelNow = mAppliedPanelTab Then Exit Sub
    mAppliedPanelTab = panelNow
    hostApp.ReapplyTabMembers mComponentId, panelNow
End Sub

Private Function SelectTab(ByVal hostApp As ReDimUI, ByVal tabNumber As Long) As Boolean
    If tabNumber < 1 Or tabNumber > ItemCount Then Exit Function
    If tabNumber = ShownTab() Then Exit Function
    mValue = tabNumber
    If LenB(mWritesToKey) > 0 Then
        hostApp.SetState mWritesToKey, ItemStateValue(ItemAt(tabNumber))
    End If
    Touch
    SelectTab = True
End Function

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
        ":" & PointerKey() & JoinedSig(mItems)) Then Exit Sub
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
        ElseIf PointerOn("row" & idx) Then
            ring.Line.ForeColor.RGB = themeValue.PrimaryColor
            ring.Line.Weight = 1.25
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
        DropPart hostSheet, "c" & idx
        DropPart hostSheet, "d" & idx
        DropPart hostSheet, "t" & idx
    Next idx
    mAppliedOptionCount = rowTotal
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
    Dim shownAt() As Long
    Dim shownTotal As Long
    Dim cursorAt As Long
    Dim entry As Variant

    EnsureCheckedSet
    If Not mItems Is Nothing Then rowTotal = mItems.Count
    If rowTotal = 0 Then
        SyncKeyCursor hostSheet, themeValue, False, 0, 0, 0, 0
        Exit Sub
    End If
    PruneTransferSelection mCheckedSet, rowTotal
    ' A filter lays the rows out anew: it is part of the layout.
    layoutKey = PartsState(themeValue) & ":" & CStr(mSelectAllRow) & _
        ":" & CStr(rowTotal) & Chr$(1) & mFilterCheck
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
    ' Each item's row under the filter, 0 when hidden; the header counts
    ' the rows that show.
    ReDim shownAt(1 To rowTotal)
    For idx = 1 To rowTotal
        If LenB(mFilterCheck) = 0 Then
            shownAt(idx) = idx
        ElseIf InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            shownTotal = shownTotal + 1
            shownAt(idx) = shownTotal
        End If
    Next idx
    If LenB(mFilterCheck) = 0 Then
        shownTotal = rowTotal
        checkedTotal = mCheckedSet.Count
    Else
        For idx = 1 To rowTotal
            If shownAt(idx) > 0 And checkedFlags(idx) Then checkedTotal = checkedTotal + 1
        Next idx
    End If
    headerRows = 0
    If mSelectAllRow Then headerRows = 1
    rowPitch = mHeight / (rowTotal + headerRows)
    boxSize = rowPitch - 6
    If boxSize > 14 Then boxSize = 14
    If boxSize < 10 Then boxSize = 10

    If mSelectAllRow Then
        If LenB(mFilterCheck) = 0 Then
            headerText = Phrase(UI_SELECT_ALL, checkedTotal, shownTotal)
        Else
            headerText = Phrase(UI_SELECT_ALL_FILTERED, mFilterCheck, checkedTotal, shownTotal)
        End If
        headerKey = CStr(checkedTotal) & "/" & CStr(shownTotal) & _
            IIf(PointerOn("row0"), "h", vbNullString) & Chr$(1) & headerText
        priorKey = mAppliedCheckHeaderSig
        If headerKey <> priorKey Then
            mAppliedCheckHeaderSig = headerKey
            Set box = EnsurePartShape( _
                hostSheet, "mb", msoShapeRoundedRectangle, boxCreated)
            Set captionShape = EnsurePartShape( _
                hostSheet, "mt", msoShapeRectangle, captionCreated)
            If LenB(priorKey) > 0 And Not (boxCreated Or captionCreated) Then
                StyleCheckBoxState box, boxSize, themeValue, _
                    (checkedTotal = shownTotal And shownTotal > 0), (checkedTotal > 0), _
                    PointerOn("row0")
                captionShape.TextFrame2.TextRange.Text = headerText
            Else
                On Error Resume Next
                box.Adjustments(1) = 0.2
                On Error GoTo 0
                StyleCheckBoxFace box, boxSize, mTop, rowPitch, themeValue, _
                    (checkedTotal = shownTotal And shownTotal > 0), (checkedTotal > 0), _
                    PointerOn("row0")
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
        DropPart hostSheet, "mb"
        DropPart hostSheet, "mt"
    End If

    For idx = 1 To rowTotal
        ' The first character carries the box state: checked adds 1 and
        ' the pointer adds 2. A row the filter hides keys as "~".
        If shownAt(idx) = 0 Then
            rowKey = "~"
        Else
            rowKey = CStr(IIf(checkedFlags(idx), 1, 0) + _
                IIf(PointerOn("row" & idx), 2, 0)) & ItemAt(idx)
        End If
        priorKey = mAppliedRowStates(idx)
        If rowKey = "~" And priorKey <> "~" Then
            mAppliedRowStates(idx) = rowKey
            On Error Resume Next
            If idx = 1 Then
                mainShape.Visible = msoFalse
            Else
                hostSheet.Shapes(ShapeName & PART_SEPARATOR & "b" & idx).Visible = msoFalse
            End If
            hostSheet.Shapes(ShapeName & PART_SEPARATOR & "t" & idx).Visible = msoFalse
            On Error GoTo 0
        ElseIf rowKey <> priorKey And rowKey <> "~" Then
            mAppliedRowStates(idx) = rowKey
            rowTop = mTop + (shownAt(idx) - 1 + headerRows) * rowPitch
            boxCreated = False
            If idx = 1 Then
                Set box = mainShape
            Else
                Set box = EnsurePartShape( _
                    hostSheet, "b" & idx, msoShapeRoundedRectangle, boxCreated)
            End If
            Set captionShape = EnsurePartShape( _
                hostSheet, "t" & idx, msoShapeRectangle, captionCreated)
            If LenB(priorKey) > 0 And priorKey <> "~" _
                And Not (boxCreated Or captionCreated) Then
                ' Drawn under this layout: rewrite only the box state or the
                ' caption text, whichever moved.
                If Left$(priorKey, 1) <> Left$(rowKey, 1) Then
                    StyleCheckBoxState box, boxSize, themeValue, _
                        checkedFlags(idx), checkedFlags(idx), PointerOn("row" & idx)
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
                    checkedFlags(idx), checkedFlags(idx), PointerOn("row" & idx)
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
            DropPart hostSheet, "b" & idx
            DropPart hostSheet, "t" & idx
        Next idx
        mAppliedOptionCount = rowTotal
    End If
    If mKeyRow > rowTotal Then mKeyRow = rowTotal
    If mKeyRow < 1 - headerRows Then mKeyRow = 1
    cursorAt = 0
    If mKeyRow >= 1 Then cursorAt = shownAt(mKeyRow)
    SyncKeyCursor hostSheet, themeValue, _
        ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId) And mVisible _
            And (mKeyRow = 0 Or cursorAt > 0), _
        mLeft - 2, mTop + (cursorAt - 1 + headerRows) * rowPitch, _
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
    ByVal partialState As Boolean, _
    Optional ByVal pointerHere As Boolean = False _
)
    StyleCheckBoxEdge box, themeValue, partialState, pointerHere
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
    ByVal partialState As Boolean, _
    Optional ByVal pointerHere As Boolean = False _
)
    box.Left = mLeft
    box.Top = rowTop + (rowPitch - boxSize) / 2
    box.Width = boxSize
    box.Height = boxSize
    box.Fill.Visible = msoTrue
    StyleCheckBoxEdge box, themeValue, partialState, pointerHere
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

' A check box's fill and edge: the accent when checked or mixed, else
' the surface with a border, which takes the accent under the pointer.
Private Sub StyleCheckBoxEdge( _
    ByVal box As Shape, _
    ByVal themeValue As ReDimUI, _
    ByVal filledState As Boolean, _
    ByVal pointerHere As Boolean _
)
    If filledState And mEnabled Then
        If pointerHere Then
            box.Fill.ForeColor.RGB = BlendRgb(ResolveFill(themeValue), _
                themeValue.OnPrimaryColor, 0.08)
        Else
            box.Fill.ForeColor.RGB = ResolveFill(themeValue)
        End If
        box.Line.Visible = msoFalse
    Else
        box.Fill.ForeColor.RGB = themeValue.SurfaceColor
        box.Line.Visible = msoTrue
        box.Line.ForeColor.RGB = _
            IIf(pointerHere, themeValue.PrimaryColor, themeValue.BorderColor)
        box.Line.Weight = 1.25
    End If
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

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mValue) & _
        ":" & PointerKey()) Then Exit Sub
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
    minusButton.Fill.ForeColor.RGB = PointerTint(themeValue.MutedColor, _
        ResolveDrawnInk(themeValue), "minus")
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
    plusButton.Fill.ForeColor.RGB = PointerTint(themeValue.MutedColor, _
        ResolveDrawnInk(themeValue), "plus")
    plusButton.Line.Visible = msoFalse
    StyleGlyphText plusButton, "+", ResolveDrawnInk(themeValue), _
        themeValue.BaseFontSize + 2, themeValue
    plusButton.Visible = BoolToMso(mVisible)
    plusButton.OnAction = mAppliedOnAction
End Sub

Private Sub ApplyToggleKnob( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim knob As Shape
    Dim knobSize As Double
    Dim targetLeft As Double
    Dim glideBase As String
    Dim knobMade As Boolean

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mChecked)) Then Exit Sub
    Set knob = EnsurePartShape(hostSheet, "knob", msoShapeOval, knobMade)
    knobSize = mHeight - 4
    If knobSize < 4 Then knobSize = 4
    knob.Height = knobSize
    knob.Top = mTop + 2
    If mChecked Then
        targetLeft = mLeft + mWidth - knobSize - 2
    Else
        targetLeft = mLeft + 2
    End If
    ' A flip of a switch drawn in place already glides the knob across.
    glideBase = PartsState(themeValue)
    PlaceGliding knob, "knob", targetLeft, knobSize, _
        Not knobMade And mAppliedValid And mAppliedChecked <> mChecked _
        And glideBase = mAppliedGlideBase
    mAppliedGlideBase = glideBase
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

' Puts a part at its target left and width, or, when it may glide and
' motion is on, leaves it where it is drawn and lets the frames close
' the distance. A part already gliding to this target keeps gliding; a
' new target takes it on from wherever it has got to.
Private Sub PlaceGliding( _
    ByVal partShape As Shape, _
    ByVal partName As String, _
    ByVal targetLeft As Double, _
    ByVal targetWidth As Double, _
    ByVal glides As Boolean _
)
    Dim leftGap As Double
    Dim widthGap As Double
    Dim alreadyGliding As Boolean

    alreadyGliding = (mGlidePart = partName)
    If alreadyGliding Then
        If Abs(mGlideTargetLeft - targetLeft) < 0.01 _
            And Abs(mGlideTargetWidth - targetWidth) < 0.01 Then Exit Sub
    End If
    If glides Or alreadyGliding Then
        If Not ReDimUI.MotionReduced Then
            leftGap = partShape.Left - targetLeft
            widthGap = partShape.Width - targetWidth
        End If
    End If
    If Abs(leftGap) < 0.75 And Abs(widthGap) < 0.75 Then
        partShape.Left = targetLeft
        partShape.Width = targetWidth
        If mGlidePart = partName Then mGlidePart = vbNullString
        Exit Sub
    End If
    ' A switch or tab strip ticks only while it glides.
    If Not alreadyGliding Then OwnerApp.TickablesChanged
    mGlidePart = partName
    mGlideTargetLeft = targetLeft
    mGlideTargetWidth = targetWidth
    mGlideGapLeft = leftGap
    mGlideGapWidth = widthGap
    ReDimUI.RequestPump
End Sub

' One frame of a glide: the distance shrinks by the time passed, the part
' moves, and under three quarters of a point it lands.
Private Sub TickGlide(ByVal hostSheet As Worksheet, ByVal elapsedMs As Long)
    Dim keepShare As Double
    Dim partName As String

    partName = mGlidePart
    keepShare = GLIDE_DECAY ^ (elapsedMs / 50#)
    mGlideGapLeft = mGlideGapLeft * keepShare
    mGlideGapWidth = mGlideGapWidth * keepShare
    If Abs(mGlideGapLeft) < 0.75 Then mGlideGapLeft = 0
    If Abs(mGlideGapWidth) < 0.75 Then mGlideGapWidth = 0
    ' Landed: cleared before the write, so a part gone missing cannot
    ' keep the frames coming, and the control leaves the ticking list.
    If mGlideGapLeft = 0 And mGlideGapWidth = 0 Then
        mGlidePart = vbNullString
        OwnerApp.TickablesChanged
    End If
    With hostSheet.Shapes(ShapeName & PART_SEPARATOR & partName)
        .Left = mGlideTargetLeft + mGlideGapLeft
        .Width = mGlideTargetWidth + mGlideGapWidth
    End With
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
    If mLeftOffset > TransferShownCount(0) - TransferRowCapacity() Then
        mLeftOffset = TransferShownCount(0) - TransferRowCapacity()
    End If
    If mLeftOffset < 0 Then mLeftOffset = 0
    If mRightOffset > TransferShownCount(1) - TransferRowCapacity() Then
        mRightOffset = TransferShownCount(1) - TransferRowCapacity()
    End If
    If mRightOffset < 0 Then mRightOffset = 0
    TransferLayout panelW, panelTop, panelH, rightX, btnX, btnTop

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

        ApplyTransferButton hostSheet, themeValue, "mvr", ">", _
            btnX, btnTop, TRANSFER_BUTTON_W, TRANSFER_BUTTON_H
        ApplyTransferButton hostSheet, themeValue, "mvar", ">>", _
            btnX, btnTop + (TRANSFER_BUTTON_H + TRANSFER_GAP), _
            TRANSFER_BUTTON_W, TRANSFER_BUTTON_H
        ApplyTransferButton hostSheet, themeValue, "mvl", "<", _
            btnX, btnTop + 2 * (TRANSFER_BUTTON_H + TRANSFER_GAP), _
            TRANSFER_BUTTON_W, TRANSFER_BUTTON_H
        ApplyTransferButton hostSheet, themeValue, "mval", "<<", _
            btnX, btnTop + 3 * (TRANSFER_BUTTON_H + TRANSFER_GAP), _
            TRANSFER_BUTTON_W, TRANSFER_BUTTON_H
        mAppliedPartTint = vbNullString
        mAppliedDropPanel = 0
    End If
    SyncTransferDropTarget mainShape, hostSheet, themeValue

    ApplyTransferHeader hostSheet, themeValue, "hl", mLeft, panelW, TRANSFER_HEADER_H, _
        TransferHeaderText(0), mAppliedHeaderLeftSig
    ApplyTransferHeader hostSheet, themeValue, "hr", rightX, panelW, TRANSFER_HEADER_H, _
        TransferHeaderText(1), mAppliedHeaderRightSig
    SyncTransferRows hostSheet, themeValue, 0, mItems, "al", _
        mLeft, panelTop, panelW, panelH, mSelLeft, mLeftOffset, _
        mAppliedLeftRows, mAppliedLeftRowStates
    SyncTransferRows hostSheet, themeValue, 1, mChosenItems, "cl", _
        rightX, panelTop, panelW, panelH, mSelRight, mRightOffset, _
        mAppliedRightRows, mAppliedRightRowStates
    ApplyTransferScroller hostSheet, themeValue, "alu", "ald", _
        mLeft, panelTop, panelW, panelH, TransferShownCount(0), mAppliedScrollLeftSig
    ApplyTransferScroller hostSheet, themeValue, "clu", "cld", _
        rightX, panelTop, panelW, panelH, TransferShownCount(1), _
        mAppliedScrollRightSig
    SyncReorderButtons hostSheet, themeValue, rightX, panelW
    SyncTransferPartTint hostSheet, themeValue

    ' The keyboard cursor, while the list holds focus.
    Dim cursorItems As Long
    Dim cursorRow As Long
    Dim cursorWidth As Double
    If mKeyRow > TransferPanelCount(mKeyPanel) Then mKeyRow = TransferPanelCount(mKeyPanel)
    cursorItems = TransferShownCount(mKeyPanel)
    cursorRow = TransferShownPosition(mKeyPanel, mKeyRow)
    If cursorRow > 0 Then cursorRow = cursorRow - IIf(mKeyPanel = 0, mLeftOffset, mRightOffset)
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
        DropPart hostSheet, upPart
        DropPart hostSheet, downPart
    End If
End Sub

' The chosen panel's reorder arrows, square pointer targets at the right
' end of its header.
Private Sub SyncReorderButtons( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal rightX As Double, _
    ByVal panelW As Double _
)
    Dim reorderKey As String
    Dim arrowShape As Shape
    Dim partNo As Long

    If mReorderable And mVisible Then
        reorderKey = PartsState(themeValue) & ":" & CStr(rightX) & ":" & CStr(panelW)
    End If
    If reorderKey = mAppliedReorderSig Then Exit Sub
    mAppliedReorderSig = reorderKey
    If LenB(reorderKey) = 0 Then
        DropPart hostSheet, "mvu"
        DropPart hostSheet, "mvd"
        Exit Sub
    End If
    For partNo = 0 To 1
        Set arrowShape = EnsurePartShape(hostSheet, IIf(partNo = 0, "mvu", "mvd"), _
            msoShapeRoundedRectangle)
        arrowShape.Width = PAGER_TARGET
        arrowShape.Height = PAGER_TARGET
        arrowShape.Left = ReorderArrowLeft(rightX, panelW, partNo)
        arrowShape.Top = mTop - 1
        arrowShape.Fill.Visible = msoTrue
        arrowShape.Fill.ForeColor.RGB = themeValue.MutedColor
        arrowShape.Line.Visible = msoFalse
        StyleGlyphText arrowShape, IIf(partNo = 0, ChrW(9650), ChrW(9660)), _
            ResolveDrawnInk(themeValue), 9, themeValue
        arrowShape.Visible = msoTrue
        arrowShape.OnAction = mAppliedOnAction
        arrowShape.ZOrder msoBringToFront
    Next partNo
    mAppliedPartTint = vbNullString
End Sub

' The left edge of the up (0) or down (1) reorder arrow.
Private Function ReorderArrowLeft( _
    ByVal rightX As Double, _
    ByVal panelW As Double, _
    ByVal partNo As Long _
) As Double
    ReorderArrowLeft = rightX + panelW - (2 - partNo) * (PAGER_TARGET + 2)
End Function

' Moves the chosen side's selected rows one place up or down among the
' rows the panel shows, a block at a time, stopping at the ends. With
' nothing selected the keys move the cursor's row. The cursor goes with
' its row.
Private Sub ReorderChosen(ByVal hostApp As ReDimUI, ByVal upward As Boolean)
    Dim itemTotal As Long
    Dim texts() As String
    Dim marked() As Boolean
    Dim shownRows() As Long
    Dim shownTotal As Long
    Dim entry As Variant
    Dim idx As Long
    Dim upper As Long
    Dim lower As Long
    Dim cursorOnly As Boolean
    Dim moved As Boolean
    Dim rebuilt As Collection

    EnsureTransferCollections
    itemTotal = mChosenItems.Count
    If itemTotal < 2 Then Exit Sub
    ReDim texts(1 To itemTotal)
    ReDim marked(1 To itemTotal)
    For Each entry In mChosenItems
        idx = idx + 1
        texts(idx) = CStr(entry)
    Next entry
    For Each entry In mSelRight
        If CLng(entry) >= 1 And CLng(entry) <= itemTotal Then marked(CLng(entry)) = True
    Next entry
    If mSelRight.Count = 0 Then
        If mKeyPanel <> 1 Or mKeyRow < 1 Or mKeyRow > itemTotal Then Exit Sub
        marked(mKeyRow) = True
        cursorOnly = True
    End If
    shownTotal = TransferShownRows(1, shownRows)
    If shownTotal < 2 Then Exit Sub
    If upward Then
        For idx = 2 To shownTotal
            upper = shownRows(idx - 1)
            lower = shownRows(idx)
            If marked(lower) And Not marked(upper) Then
                SwapChosenRows texts, marked, upper, lower
                moved = True
            End If
        Next idx
    Else
        For idx = shownTotal - 1 To 1 Step -1
            upper = shownRows(idx)
            lower = shownRows(idx + 1)
            If marked(upper) And Not marked(lower) Then
                SwapChosenRows texts, marked, upper, lower
                moved = True
            End If
        Next idx
    End If
    If Not moved Then Exit Sub
    Set rebuilt = New Collection
    For idx = 1 To itemTotal
        rebuilt.Add texts(idx)
    Next idx
    Set mChosenItems = rebuilt
    Set mSelRight = New Collection
    For idx = 1 To itemTotal
        If marked(idx) And Not cursorOnly Then mSelRight.Add idx
    Next idx
    Touch
    CommitTransferChange hostApp
End Sub

' Two chosen rows trade places, marks and the key cursor with them.
Private Sub SwapChosenRows( _
    ByRef texts() As String, _
    ByRef marked() As Boolean, _
    ByVal upper As Long, _
    ByVal lower As Long _
)
    Dim heldText As String
    Dim heldMark As Boolean

    heldText = texts(upper)
    texts(upper) = texts(lower)
    texts(lower) = heldText
    heldMark = marked(upper)
    marked(upper) = marked(lower)
    marked(lower) = heldMark
    If mKeyPanel = 1 Then
        If mKeyRow = upper Then
            mKeyRow = lower
        ElseIf mKeyRow = lower Then
            mKeyRow = upper
        End If
    End If
End Sub

' The item numbers a transfer panel shows, in order, into shownItems;
' returns how many.
Private Function TransferShownRows(ByVal panelIndex As Long, ByRef shownItems() As Long) As Long
    Dim shownTotal As Long
    Dim idx As Long

    shownTotal = TransferShownCount(panelIndex)
    If shownTotal = 0 Then Exit Function
    ReDim shownItems(1 To shownTotal)
    For idx = 1 To shownTotal
        shownItems(idx) = TransferShownItem(panelIndex, idx)
    Next idx
    TransferShownRows = shownTotal
End Function

Private Function PanelFilter(ByVal panelIndex As Long) As String
    If panelIndex = 0 Then
        PanelFilter = mFilterLeft
    Else
        PanelFilter = mFilterRight
    End If
End Function

' How many rows a transfer panel shows: all its items, or with a filter
' those whose text holds it.
Private Function TransferShownCount(ByVal panelIndex As Long) As Long
    If LenB(PanelFilter(panelIndex)) = 0 Then
        TransferShownCount = TransferPanelCount(panelIndex)
        Exit Function
    End If
    EnsureShownRows panelIndex
    If panelIndex = 0 Then
        TransferShownCount = mShownLeftCount
    Else
        TransferShownCount = mShownRightCount
    End If
End Function

' The item number at a shown position, 0 past the end.
Private Function TransferShownItem(ByVal panelIndex As Long, ByVal shownPosition As Long) As Long
    If shownPosition < 1 Then Exit Function
    If LenB(PanelFilter(panelIndex)) = 0 Then
        If shownPosition <= TransferPanelCount(panelIndex) Then
            TransferShownItem = shownPosition
        End If
        Exit Function
    End If
    EnsureShownRows panelIndex
    If panelIndex = 0 Then
        If shownPosition <= mShownLeftCount Then TransferShownItem = mShownLeft(shownPosition)
    Else
        If shownPosition <= mShownRightCount Then TransferShownItem = mShownRight(shownPosition)
    End If
End Function

' The shown position of an item, 0 when the filter hides it.
Private Function TransferShownPosition(ByVal panelIndex As Long, ByVal itemIndex As Long) As Long
    Dim shownTotal As Long
    Dim idx As Long

    If itemIndex < 1 Then Exit Function
    If LenB(PanelFilter(panelIndex)) = 0 Then
        If itemIndex <= TransferPanelCount(panelIndex) Then
            TransferShownPosition = itemIndex
        End If
        Exit Function
    End If
    shownTotal = TransferShownCount(panelIndex)
    For idx = 1 To shownTotal
        If TransferShownItem(panelIndex, idx) = itemIndex Then
            TransferShownPosition = idx
            Exit Function
        End If
    Next idx
End Function

' A filtered panel's item numbers, read once per filter, item version,
' and item list.
Private Sub EnsureShownRows(ByVal panelIndex As Long)
    Dim panelItems As Collection
    Dim cacheKey As String
    Dim filterText As String
    Dim shownItems() As Long
    Dim shownTotal As Long
    Dim entry As Variant
    Dim idx As Long

    EnsureTransferCollections
    If panelIndex = 0 Then
        Set panelItems = mItems
    Else
        Set panelItems = mChosenItems
    End If
    filterText = PanelFilter(panelIndex)
    cacheKey = filterText & Chr$(1) & CStr(mItemsVersion) & ":" & _
        CStr(ObjPtr(panelItems)) & ":" & CStr(panelItems.Count)
    If panelIndex = 0 Then
        If cacheKey = mShownLeftKey Then Exit Sub
    ElseIf cacheKey = mShownRightKey Then
        Exit Sub
    End If
    ReDim shownItems(0 To panelItems.Count)
    For Each entry In panelItems
        idx = idx + 1
        If InStr(1, CStr(entry), filterText, vbTextCompare) > 0 Then
            shownTotal = shownTotal + 1
            shownItems(shownTotal) = idx
        End If
    Next entry
    If panelIndex = 0 Then
        mShownLeft = shownItems
        mShownLeftCount = shownTotal
        mShownLeftKey = cacheKey
    Else
        mShownRight = shownItems
        mShownRightCount = shownTotal
        mShownRightKey = cacheKey
    End If
End Sub

' A panel's header: its caption and count, and with a filter the
' filter in quotes and how many of the items show.
Private Function TransferHeaderText(ByVal panelIndex As Long) As String
    Dim captionText As String

    If panelIndex = 0 Then
        captionText = mCaptionLeft
        If LenB(captionText) = 0 Then captionText = Phrase(UI_AVAILABLE)
    Else
        captionText = mCaptionRight
        If LenB(captionText) = 0 Then captionText = Phrase(UI_SELECTED)
    End If
    If LenB(PanelFilter(panelIndex)) = 0 Then
        TransferHeaderText = captionText & " (" & TransferPanelCount(panelIndex) & ")"
    Else
        TransferHeaderText = Phrase(UI_PANEL_FILTERED, captionText, PanelFilter(panelIndex), _
            TransferShownCount(panelIndex), TransferPanelCount(panelIndex))
    End If
End Function

' Typing while a transfer or check list has the keys builds a filter:
' a character adds to it, Backspace takes one off. True when the key
' was the filter's. The transfer list filters the panel its cursor is
' in, and the cursor moves to the first row left showing.
Private Function TypeListFilter(ByVal keyText As String) As Boolean
    Dim filterText As String

    Select Case keyText
        Case "{APOS}"
            keyText = "'"
        Case "{QUOTE}"
            keyText = """"
    End Select
    If mKind = KIND_TRANSFER Then
        filterText = PanelFilter(mKeyPanel)
    Else
        filterText = mFilterCheck
    End If
    If keyText = "{BS}" Then
        If LenB(filterText) = 0 Then Exit Function
        filterText = Left$(filterText, Len(filterText) - 1)
    ElseIf Len(keyText) = 1 And keyText <> " " Then
        filterText = filterText & keyText
    Else
        Exit Function
    End If
    SetListFilter filterText
    TypeListFilter = True
End Function

' Sets the filter of the transfer panel the cursor is in, or of the
' check list, and rewinds to its first row.
Private Sub SetListFilter(ByVal filterText As String)
    If mKind = KIND_TRANSFER Then
        If mKeyPanel = 0 Then
            mFilterLeft = filterText
            mLeftOffset = 0
        Else
            mFilterRight = filterText
            mRightOffset = 0
        End If
        PlaceKeyCursor mKeyPanel, TransferShownItem(mKeyPanel, 1)
    Else
        mFilterCheck = filterText
        PlaceKeyCursor 0, CheckShownItem(1)
    End If
    Touch
End Sub

' Esc on a filtered list clears the filter first. True when it did.
Private Function ClearListFilter() As Boolean
    If mKind = KIND_TRANSFER Then
        If LenB(PanelFilter(mKeyPanel)) = 0 Then Exit Function
    ElseIf mKind = KIND_CHECKLIST Then
        If LenB(mFilterCheck) = 0 Then Exit Function
    ElseIf mKind = KIND_TABLE Then
        If LenB(mTableFilter) = 0 Then Exit Function
        SetTableFilter vbNullString
        ClearListFilter = True
        Exit Function
    Else
        Exit Function
    End If
    SetListFilter vbNullString
    ClearListFilter = True
End Function

' The item a check list shows at a row position, counting from 1 below
' any select-all row; 0 past the end.
Private Function CheckShownItem(ByVal shownPosition As Long) As Long
    Dim idx As Long
    Dim seen As Long

    If shownPosition < 1 Then Exit Function
    If LenB(mFilterCheck) = 0 Then
        If shownPosition <= ItemCount Then CheckShownItem = shownPosition
        Exit Function
    End If
    For idx = 1 To ItemCount
        If InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            seen = seen + 1
            If seen = shownPosition Then
                CheckShownItem = idx
                Exit Function
            End If
        End If
    Next idx
End Function

' How many rows a check list shows under its filter.
Private Function CheckShownCount() As Long
    Dim idx As Long

    If LenB(mFilterCheck) = 0 Then
        CheckShownCount = ItemCount
        Exit Function
    End If
    For idx = 1 To ItemCount
        If InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            CheckShownCount = CheckShownCount + 1
        End If
    Next idx
End Function

' A check list item's row position under the filter, 0 when hidden.
Private Function CheckShownPosition(ByVal itemIndex As Long) As Long
    Dim idx As Long

    If itemIndex < 1 Or itemIndex > ItemCount Then Exit Function
    If LenB(mFilterCheck) = 0 Then
        CheckShownPosition = itemIndex
        Exit Function
    End If
    If InStr(1, ItemAt(itemIndex), mFilterCheck, vbTextCompare) = 0 Then Exit Function
    For idx = 1 To itemIndex
        If InStr(1, ItemAt(idx), mFilterCheck, vbTextCompare) > 0 Then
            CheckShownPosition = CheckShownPosition + 1
        End If
    Next idx
End Function

' The move button or paging arrow under the pointer takes the pointer
' tint; the one that had it goes back to its plain fill.
Private Sub SyncTransferPartTint(ByVal hostSheet As Worksheet, ByVal themeValue As ReDimUI)
    Dim tintKey As String
    Dim priorPart As String

    Select Case mHoveredPart
        Case "mvr", "mvar", "mvl", "mval", "alu", "ald", "clu", "cld", _
            "mvu", "mvd"
            If PointerOn(mHoveredPart) Then tintKey = PointerKey()
    End Select
    If tintKey = mAppliedPartTint Then Exit Sub
    On Error Resume Next
    If LenB(mAppliedPartTint) > 0 Then
        priorPart = Left$(mAppliedPartTint, Len(mAppliedPartTint) - 1)
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & priorPart) _
            .Fill.ForeColor.RGB = themeValue.MutedColor
    End If
    If LenB(tintKey) > 0 Then
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & mHoveredPart) _
            .Fill.ForeColor.RGB = PointerTint(themeValue.MutedColor, _
                ResolveDrawnInk(themeValue), mHoveredPart)
    End If
    On Error GoTo 0
    mAppliedPartTint = tintKey
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
    ByVal panelIndex As Long, _
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
    Dim shownTotal As Long

    shownTotal = TransferShownCount(panelIndex)
    wanted = shownTotal - scrollOffset
    If wanted > Int((panelH - 6) / TRANSFER_ROW_H) Then
        wanted = Int((panelH - 6) / TRANSFER_ROW_H)
    End If
    If wanted < 0 Then wanted = 0
    If Not mVisible Then wanted = 0
    rowWidth = panelW - 6
    If shownTotal > Int((panelH - 6) / TRANSFER_ROW_H) Then
        ' Leave room for the paging arrows.
        rowWidth = rowWidth - PAGER_TARGET - 2
    End If
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    layoutKey = PartsState(themeValue) & ":" & CStr(panelX) & ":" & _
        CStr(panelTop) & ":" & CStr(rowWidth)
    EnsureRowStates rowStates, wanted
    For idx = 1 To wanted
        itemIndex = TransferShownItem(panelIndex, scrollOffset + idx)
        rowSelected = TransferSelectionHas(selectionSet, itemIndex)
        If rowSelected Then
            rowFill = themeValue.PrimaryColor
            rowInk = themeValue.OnPrimaryColor
        Else
            rowFill = ResolveSurfaceFill(themeValue)
            rowInk = themeValue.OnSurfaceColor
        End If
        rowFill = PointerTint(rowFill, rowInk, partPrefix & idx)
        rowText = GutterText(rowSelected, CStr(itemsRef.Item(itemIndex)))
        rowKey = layoutKey & Chr$(1) & CStr(rowFill) & ":" & _
            CStr(rowInk) & ":" & rowText
        priorKey = rowStates(idx)
        If rowKey <> priorKey Then
            rowStates(idx) = rowKey
            Set rowShape = EnsurePartShape( _
                hostSheet, partPrefix & idx, msoShapeRoundedRectangle, partCreated)
            If KeyFrame(priorKey) = layoutKey And Not partCreated Then
                ' Same frame: only the fill, the text, and the ink move.
                rowShape.Fill.ForeColor.RGB = rowFill
                With rowShape.TextFrame2.TextRange
                    .Text = rowText
                    .Font.Fill.ForeColor.RGB = rowInk
                End With
            Else
                ' The hover and selection fills round their corners inside
                ' the rounded panel.
                RoundCorners rowShape, TRANSFER_ROW_H
                rowShape.Left = panelX + 3
                rowShape.Top = panelTop + 3 + (idx - 1) * TRANSFER_ROW_H
                rowShape.Width = rowWidth
                rowShape.Height = TRANSFER_ROW_H
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
        DropPart hostSheet, partPrefix & idx
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

    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(mStyleVariant)) Then Exit Sub
    Set caret = EnsurePartShape(hostSheet, "caret", msoShapeIsoscelesTriangle)
    caret.Width = 8
    caret.Height = 5
    caret.Left = mLeft + mWidth - rightInset
    caret.Top = mTop + (mHeight - 5) / 2
    caret.Rotation = 180
    caret.Fill.Visible = msoTrue
    If mIsMenu Then
        caret.Fill.ForeColor.RGB = ResolveTextColor(themeValue)
    Else
        caret.Fill.ForeColor.RGB = themeValue.OnMutedColor
    End If
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
    Dim bubbleOn As Boolean

    fraction = SlideFraction()
    bubbleOn = mVisible And (mSlideTracking Or _
        ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId))
    If ExtrasCurrent(PartsState(themeValue) & ":" & CStr(fraction) & _
        ":" & CStr(mSlideTracking) & ":" & CStr(bubbleOn) & ":" & _
        PointerKey()) Then Exit Sub
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
        thumb.Fill.ForeColor.RGB = PointerTint(RGB(255, 255, 255), _
            themeValue.OnSurfaceColor)
    End If
    thumb.Line.Visible = msoTrue
    thumb.Line.ForeColor.RGB = themeValue.BorderColor
    thumb.Line.Weight = 1.25
    thumb.Visible = BoolToMso(mVisible)
    thumb.OnAction = mAppliedOnAction
    SyncSlideBubble hostSheet, themeValue, bubbleOn, _
        thumb.Left + thumbSize / 2
End Sub

' The value in a bubble over the thumb while a drag or keyboard focus
' moves it; below the track when there is no room above.
Private Sub SyncSlideBubble( _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI, _
    ByVal bubbleOn As Boolean, _
    ByVal centerX As Double _
)
    Dim bubble As Shape
    Dim bubbleText As String
    Dim bubbleSize As Double
    Dim bubbleWidth As Double

    If Not bubbleOn Then
        If mAppliedBubble Then
            mAppliedBubble = False
            DropPart hostSheet, "vb"
        End If
        Exit Sub
    End If
    bubbleText = CStr(mValue)
    bubbleSize = mFontSize
    If bubbleSize <= 0 Then bubbleSize = themeValue.BaseFontSize
    bubbleSize = bubbleSize - 1
    bubbleWidth = Len(bubbleText) * bubbleSize * 0.62 + 12
    If bubbleWidth < 24 Then bubbleWidth = 24
    Set bubble = EnsurePartShape(hostSheet, "vb", msoShapeRoundedRectangle)
    bubble.Width = bubbleWidth
    bubble.Height = bubbleSize + 8
    bubble.Left = centerX - bubbleWidth / 2
    If mTop - bubble.Height - 4 >= 0 Then
        bubble.Top = mTop - bubble.Height - 4
    Else
        bubble.Top = mTop + mHeight + 4
    End If
    bubble.Fill.Visible = msoTrue
    bubble.Fill.ForeColor.RGB = themeValue.PrimaryColor
    bubble.Line.Visible = msoFalse
    With bubble.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = bubbleText
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = bubbleSize
            .Bold = msoTrue
            .Fill.ForeColor.RGB = themeValue.OnPrimaryColor
        End With
    End With
    bubble.Visible = msoTrue
    bubble.OnAction = mAppliedOnAction
    mAppliedBubble = True
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
' The press watch each frame: a slider's drag, a transfer list's row
' drag.
Friend Sub TickPressDrag(ByVal hostApp As ReDimUI, ByVal buttonDown As Boolean)
    Dim fraction As Double

    If mKind = KIND_TRANSFER Then
        TickTransferDrag hostApp, buttonDown
        Exit Sub
    End If
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
        ' Sheet points on a sheet behind the active one are not under
        ' the pointer, whatever their numbers.
        If mVisible And mEnabled And SheetInFront(hostApp) Then
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

Private Function SheetInFront(ByVal hostApp As ReDimUI) As Boolean
    On Error Resume Next
    SheetInFront = (hostApp.Sheet Is ActiveSheet)
    On Error GoTo 0
End Function

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
    If mKind = KIND_DATE Then
        ListAreaHolds = (pointsX >= mLeft - 2 And pointsX <= mLeft + mWidth + 2 _
            And pointsY >= mTop - 2 And pointsY <= mTop + mHeight + 2) _
            Or (pointsX >= mLeft - 2 And pointsX <= mLeft + CAL_WIDTH + 2 _
            And pointsY >= CalendarTop() - 2 _
            And pointsY <= CalendarTop() + CAL_DEPTH + 2)
        Exit Function
    End If
    ' The face, then the rows and pagers across the span they drew.
    If pointsX >= mLeft - 2 And pointsX <= mLeft + mWidth + 2 _
        And pointsY >= mTop - 2 And pointsY <= mTop + mHeight + 2 Then
        ListAreaHolds = True
        Exit Function
    End If
    If pointsX < ListSpanLeft() - 2 Then Exit Function
    If pointsX > ListSpanLeft() + ListSpanWidth() + 2 Then Exit Function
    If pointsY < ListTop() - 2 Then Exit Function
    If pointsY > ListTop() + DropListDepth() + 2 Then Exit Function
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
    If (mKind = KIND_COMBO Or mKind = KIND_SELECT Or mKind = KIND_DATE) _
        And mSelectOpen Then
        ChromeClaimsPoint = ListAreaHolds(pointsX, pointsY)
    End If
End Function

' True when this control's open drop list holds the point: its face or
' its rows and pagers.
Friend Function ListHoldsPoint(ByVal pointsX As Double, ByVal pointsY As Double) As Boolean
    If Not mVisible Then Exit Function
    If (mKind = KIND_COMBO Or mKind = KIND_SELECT Or mKind = KIND_DATE) _
        And mSelectOpen Then
        ListHoldsPoint = ListAreaHolds(pointsX, pointsY)
    End If
End Function

' True when this visible control's rectangle holds the point.
Friend Function CoversPoint(ByVal pointsX As Double, ByVal pointsY As Double) As Boolean
    If Not mVisible Or mKind = KIND_STACK Then Exit Function
    CoversPoint = (pointsX >= mLeft And pointsX <= mLeft + mWidth _
        And pointsY >= mTop And pointsY <= mTop + mHeight)
End Function

' The part of this control under a sheet point, as the pointer watches
' name it: "" for the face; "minus" or "plus" on a stepper; "row" and a
' number for radio and check-list rows, 0 for select-all; "opt" and a
' number for an open list's rows, "optu" and "optd" for its pagers; and
' a transfer list's row, move-button, and paging-arrow parts.
Friend Function PartAtPoint(ByVal pointsX As Double, ByVal pointsY As Double) As String
    Dim slot As Long
    Dim itemTotal As Long
    Dim headerRows As Long

    If Not mItems Is Nothing Then itemTotal = mItems.Count
    Select Case mKind
        Case KIND_STEPPER
            If pointsX < mLeft + mHeight Then
                PartAtPoint = "minus"
            ElseIf pointsX > mLeft + mWidth - mHeight Then
                PartAtPoint = "plus"
            End If
        Case KIND_RADIO
            If itemTotal > 0 Then
                slot = Int((pointsY - mTop) / (mHeight / itemTotal)) + 1
                If slot >= 1 And slot <= itemTotal Then PartAtPoint = "row" & slot
            End If
        Case KIND_CHECKLIST
            If itemTotal > 0 Then
                If mSelectAllRow Then headerRows = 1
                slot = Int((pointsY - mTop) / (mHeight / (itemTotal + headerRows))) _
                    + 1 - headerRows
                If slot = 0 And headerRows = 1 Then
                    PartAtPoint = "row0"
                ElseIf slot >= 1 Then
                    ' Rows below the filter's last show nothing.
                    slot = CheckShownItem(slot)
                    If slot >= 1 Then PartAtPoint = "row" & slot
                End If
            End If
        Case KIND_SELECT, KIND_COMBO
            If mSelectOpen Then PartAtPoint = ListPartAt(pointsY)
        Case KIND_TRANSFER
            PartAtPoint = TransferPartAt(pointsX, pointsY)
        Case KIND_DATE
            If mSelectOpen Then PartAtPoint = CalendarPartAt(pointsX, pointsY)
        Case KIND_TABLE
            PartAtPoint = TablePartAt(pointsX, pointsY)
        Case KIND_TABS
            If itemTotal > 0 And mAppliedOptionCount = itemTotal Then
                For slot = 1 To itemTotal
                    If pointsX - mLeft < mTabRights(slot) Then
                        PartAtPoint = "tab" & slot
                        Exit For
                    End If
                Next slot
            End If
    End Select
End Function

' The open list's part at a height: nothing on the face, else a row or
' a pager.
Private Function ListPartAt(ByVal pointsY As Double) As String
    Dim slot As Long
    Dim rowShift As Long
    Dim wanted As Long
    Dim matchTotal As Long

    If pointsY >= mTop And pointsY <= mTop + mHeight Then Exit Function
    If mListMatches Is Nothing Then Exit Function
    matchTotal = mListMatches.Count
    wanted = matchTotal - mListOffset
    If wanted > mListRows Then wanted = mListRows
    If mListOffset > 0 Then rowShift = 1
    If pointsY < ListTop() Then Exit Function
    slot = Int((pointsY - ListTop()) / DropRowPitch())
    If rowShift = 1 And slot = 0 Then
        ListPartAt = "optu"
    ElseIf slot - rowShift + 1 <= wanted Then
        ListPartAt = "opt" & (slot - rowShift + 1)
    ElseIf slot - rowShift = wanted _
        And matchTotal - mListOffset - wanted > 0 Then
        ListPartAt = "optd"
    End If
End Function

' A transfer list's part at a point: a move button, a panel's paging
' arrow, or a row, named as the list draws them.
Private Function TransferPartAt(ByVal pointsX As Double, ByVal pointsY As Double) As String
    Dim panelW As Double
    Dim panelTop As Double
    Dim panelH As Double
    Dim rightX As Double
    Dim btnX As Double
    Dim btnTop As Double
    Dim slot As Long

    If mItems Is Nothing Or mChosenItems Is Nothing Then Exit Function
    TransferLayout panelW, panelTop, panelH, rightX, btnX, btnTop
    If mReorderable And pointsY >= mTop - 1 And pointsY < panelTop Then
        If pointsX >= ReorderArrowLeft(rightX, panelW, 1) Then
            If pointsX <= ReorderArrowLeft(rightX, panelW, 1) + PAGER_TARGET Then
                TransferPartAt = "mvd"
            End If
        ElseIf pointsX >= ReorderArrowLeft(rightX, panelW, 0) Then
            If pointsX <= ReorderArrowLeft(rightX, panelW, 0) + PAGER_TARGET Then
                TransferPartAt = "mvu"
            End If
        End If
        Exit Function
    End If
    If pointsY < panelTop Or pointsY > panelTop + panelH Then Exit Function
    If pointsX >= btnX And pointsX <= btnX + TRANSFER_BUTTON_W Then
        slot = Int((pointsY - btnTop) / (TRANSFER_BUTTON_H + TRANSFER_GAP))
        If slot < 0 Or slot > 3 Then Exit Function
        If pointsY > btnTop + slot * (TRANSFER_BUTTON_H + TRANSFER_GAP) _
            + TRANSFER_BUTTON_H Then Exit Function
        Select Case slot
            Case 0
                TransferPartAt = "mvr"
            Case 1
                TransferPartAt = "mvar"
            Case 2
                TransferPartAt = "mvl"
            Case Else
                TransferPartAt = "mval"
        End Select
    ElseIf pointsX >= mLeft And pointsX <= mLeft + panelW Then
        TransferPartAt = TransferPanelPartAt("al", mLeft, panelTop, panelW, _
            panelH, TransferShownCount(0), mLeftOffset, pointsX, pointsY)
    ElseIf pointsX >= rightX And pointsX <= rightX + panelW Then
        TransferPartAt = TransferPanelPartAt("cl", rightX, panelTop, panelW, _
            panelH, TransferShownCount(1), mRightOffset, pointsX, pointsY)
    End If
End Function

' Within one panel: its paging arrows, when the list outgrows it, then
' its rows. The arrows take the panel's prefix with "u" or "d".
Private Function TransferPanelPartAt( _
    ByVal partPrefix As String, _
    ByVal panelX As Double, _
    ByVal panelTop As Double, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal itemTotal As Long, _
    ByVal scrollOffset As Long, _
    ByVal pointsX As Double, _
    ByVal pointsY As Double _
) As String
    Dim slot As Long

    If itemTotal > TransferRowCapacity() Then
        If pointsX >= panelX + panelW - PAGER_TARGET - 3 Then
            If pointsY <= panelTop + 3 + PAGER_TARGET Then
                TransferPanelPartAt = partPrefix & "u"
            ElseIf pointsY >= panelTop + panelH - PAGER_TARGET - 3 Then
                TransferPanelPartAt = partPrefix & "d"
            End If
            Exit Function
        End If
    End If
    If pointsY < panelTop + 3 Then Exit Function
    slot = Int((pointsY - panelTop - 3) / TRANSFER_ROW_H) + 1
    If slot <= itemTotal - scrollOffset And slot <= TransferRowCapacity() Then
        TransferPanelPartAt = partPrefix & slot
    End If
End Function

' The transfer list's panels and move buttons, in points.
Private Sub TransferLayout( _
    ByRef panelW As Double, _
    ByRef panelTop As Double, _
    ByRef panelH As Double, _
    ByRef rightX As Double, _
    ByRef btnX As Double, _
    ByRef btnTop As Double _
)
    panelW = (mWidth - TRANSFER_BUTTON_W - 4 * TRANSFER_GAP) / 2
    If panelW < 40 Then panelW = 40
    panelTop = mTop + TRANSFER_HEADER_H
    panelH = mHeight - TRANSFER_HEADER_H
    If panelH < 30 Then panelH = 30
    rightX = mLeft + panelW + TRANSFER_BUTTON_W + 4 * TRANSFER_GAP
    btnX = mLeft + panelW + 2 * TRANSFER_GAP
    btnTop = panelTop + (panelH - (4 * TRANSFER_BUTTON_H + 3 * TRANSFER_GAP)) / 2
    If btnTop < panelTop Then btnTop = panelTop
End Sub

' The tooltip, in inverse colors: DisabledReason while the control is
' disabled and has one, else Tooltip. It fits its words, wrapping past
' TIP_MAX_WIDTH. A control no taller than a field row keeps its whole
' drawing clear: the tip sits below the hint or message, or above the
' caption when the window has no room below. On a taller control, a table
' or a list, the tip sits under the pointer, flipping above it or left to
' stay in view. None while the control's list or calendar is open, which
' the tip would cover. False when there is nothing to show; the tip's
' edges come back for the pointer watch.
Friend Function DrawTooltip( _
    ByVal pointsX As Double, _
    ByVal pointsY As Double, _
    ByRef tipLeft As Double, _
    ByRef tipTop As Double, _
    ByRef tipRight As Double, _
    ByRef tipBottom As Double _
) As Boolean
    Dim tipText As String
    Dim tipShape As Shape
    Dim hostApp As ReDimUI
    Dim themeValue As ReDimUI
    Dim tipWidth As Double
    Dim tipHeight As Double
    Dim spanTop As Double
    Dim spanBottom As Double
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double

    If mSelectOpen Then Exit Function
    If Not mEnabled And LenB(mDisabledReason) > 0 Then
        tipText = mDisabledReason
    Else
        ' A shortcut shows with the tip, or as the tip, as Office does.
        tipText = mTooltip
        If LenB(mShortcut) > 0 Then
            If LenB(tipText) > 0 Then
                tipText = tipText & " (" & ShortcutLabel(mShortcut) & ")"
            Else
                tipText = ShortcutLabel(mShortcut)
            End If
        End If
    End If
    If LenB(tipText) = 0 Then Exit Function
    Set hostApp = OwnerApp
    Set themeValue = hostApp.Theme
    Set tipShape = EnsurePartShape(hostApp.Sheet, "tt", msoShapeRoundedRectangle)
    tipShape.Fill.Visible = msoTrue
    tipShape.Fill.ForeColor.RGB = themeValue.OnSurfaceColor
    tipShape.Line.Visible = msoFalse
    With tipShape.TextFrame2
        .MarginLeft = 7
        .MarginRight = 7
        .MarginTop = 3
        .MarginBottom = 3
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .TextRange.Text = tipText
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = themeValue.BaseFontSize - 1
            .Bold = msoFalse
            .Fill.ForeColor.RGB = themeValue.SurfaceColor
        End With
        ' A long tip wraps. Office refits a shape to its text only when the
        ' text or the fit changes, so the fit goes off and on again for the
        ' height to follow the wrap.
        If tipShape.Width > TIP_MAX_WIDTH Then
            .WordWrap = msoTrue
            tipShape.Width = TIP_MAX_WIDTH
            .AutoSize = msoAutoSizeNone
            .AutoSize = msoAutoSizeShapeToFitText
        End If
    End With
    tipWidth = tipShape.Width
    tipHeight = tipShape.Height
    hostApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
    tipLeft = pointsX
    If tipLeft + tipWidth > viewLeft + viewWidth Then
        tipLeft = viewLeft + viewWidth - tipWidth - 4
    End If
    If tipLeft < 0 Then tipLeft = 0
    If mHeight <= TIP_BELOW_MAX_H Then
        DrawingSpan spanTop, spanBottom
        tipTop = spanBottom + 4
        If tipTop + tipHeight > viewTop + viewHeight Then
            tipTop = spanTop - tipHeight - 4
        End If
    Else
        tipTop = pointsY + 18
        If tipTop + tipHeight > viewTop + viewHeight Then
            tipTop = pointsY - tipHeight - 6
        End If
    End If
    If tipTop < 0 Then tipTop = 0
    tipShape.Left = tipLeft
    tipShape.Top = tipTop
    tipShape.Visible = msoTrue
    tipShape.OnAction = mAppliedOnAction
    tipShape.ZOrder msoBringToFront
    tipRight = tipLeft + tipWidth
    tipBottom = tipTop + tipHeight
    DrawTooltip = True
End Function

' The span a control's drawing takes, top to bottom: the control, the
' caption above it, and the hint, message, or count below it.
Private Sub DrawingSpan(ByRef spanTop As Double, ByRef spanBottom As Double)
    Dim noteSize As Double

    noteSize = mFontSize
    If noteSize <= 0 Then noteSize = OwnerApp.Theme.BaseFontSize
    spanTop = mTop
    spanBottom = mTop + mHeight
    If LenB(mAppliedCaptionSig) > 0 Then spanTop = mTop - noteSize - 7
    If LenB(mAppliedHintSig) > 0 Or LenB(mAppliedMessageSig) > 0 _
        Or LenB(mAppliedCountSig) > 0 Then spanBottom = mTop + mHeight + noteSize + 6
End Sub

' True while this control's drop list or calendar is open.
Friend Property Get ListIsOpen() As Boolean
    ListIsOpen = mSelectOpen
End Property

Friend Sub ClearTooltipPart()
    On Error Resume Next
    OwnerApp.Sheet.Shapes(ShapeName & PART_SEPARATOR & "tt").Delete
    On Error GoTo 0
End Sub

' True for a visible, enabled stepper or transfer list, whose presses the
' pump watches while its sheet is in front.
Friend Function NeedsPressWatch() As Boolean
    If mKind <> KIND_STEPPER And mKind <> KIND_TRANSFER Then Exit Function
    NeedsPressWatch = mVisible And mEnabled And Not mBusy
End Function

' The parts a held press repeats.
Friend Function RepeatsPart(ByVal partName As String) As Boolean
    If Not mVisible Or Not mEnabled Or mBusy Then Exit Function
    Select Case mKind
        Case KIND_STEPPER
            RepeatsPart = (partName = "minus" Or partName = "plus")
        Case KIND_SELECT, KIND_COMBO
            RepeatsPart = mSelectOpen And (partName = "optu" Or partName = "optd")
        Case KIND_TRANSFER
            Select Case partName
                Case "alu", "ald", "clu", "cld"
                    RepeatsPart = True
            End Select
    End Select
End Function

Friend Sub BeginHold()
    mHoldValueAtStart = mValue
End Sub

' One repeat of a held part: a stepper step, which writes its state but
' leaves OnChange to the release, or a page of a list.
Friend Sub RepeatPart(ByVal hostApp As ReDimUI, ByVal partName As String)
    Select Case partName
        Case "minus"
            SetStepperValue hostApp, mValue - mSliderStep
        Case "plus"
            SetStepperValue hostApp, mValue + mSliderStep
        Case "optu", "optd"
            PageDropList (partName = "optd")
        Case Else
            PageTransferPanel partName
    End Select
End Sub

' A hold that repeated swallows its release click and, for a stepper
' that moved, runs OnChange once.
Friend Sub EndHold(ByVal hostApp As ReDimUI, ByVal repeated As Boolean)
    If Not repeated Then Exit Sub
    mHoldSuppressUntilMs = GetTickCount64() + 400
    If mKind = KIND_STEPPER And mValue <> mHoldValueAtStart Then
        On Error GoTo HandlerFailed
        RunClickHandlers hostApp
    End If
    Exit Sub
HandlerFailed:
    hostApp.ReportAppError "ReDim handler error " & Err.Number & ": " & Err.Description
End Sub

' The second click on the same row within the Windows double-click time
' is a double click.
Private Function RowDoubleClicked(ByVal partName As String) As Boolean
    Dim nowMs As LongLong

    nowMs = GetTickCount64()
    If partName = mLastRowClickPart _
        And nowMs - mLastRowClickMs <= GetDoubleClickTime() Then
        RowDoubleClicked = True
        mLastRowClickPart = vbNullString
    Else
        mLastRowClickPart = partName
        mLastRowClickMs = nowMs
    End If
End Function

' Moves one row's item across. The panel's other selected rows stay
' selected, their numbers closing over the moved row.
Private Sub MoveTransferRow( _
    ByVal hostApp As ReDimUI, _
    ByVal rightward As Boolean, _
    ByVal itemIndex As Long _
)
    Dim priorSelection As Collection
    Dim keptSelection As Collection
    Dim soleRow As Collection
    Dim entry As Variant

    EnsureTransferCollections
    If rightward Then
        Set priorSelection = mSelLeft
    Else
        Set priorSelection = mSelRight
    End If
    Set keptSelection = New Collection
    For Each entry In priorSelection
        If CLng(entry) < itemIndex Then
            keptSelection.Add CLng(entry)
        ElseIf CLng(entry) > itemIndex Then
            keptSelection.Add CLng(entry) - 1
        End If
    Next entry
    Set soleRow = New Collection
    soleRow.Add itemIndex
    If rightward Then
        Set mSelLeft = soleRow
    Else
        Set mSelRight = soleRow
    End If
    TransferSelected hostApp, rightward
    If rightward Then
        Set mSelLeft = keptSelection
    Else
        Set mSelRight = keptSelection
    End If
    Touch
End Sub

' A press on a transfer row, held and moved: along its panel it selects
' the rows from the pressed one to the one under the pointer; onto the
' other panel it outlines that panel, and the release there moves the
' pressed row, or the whole selection when the row is part of it. The
' release click after either is swallowed.
Private Sub TickTransferDrag(ByVal hostApp As ReDimUI, ByVal buttonDown As Boolean)
    Dim pointsX As Double
    Dim pointsY As Double
    Dim panelAt As Long
    Dim rowAt As Long
    Dim rowSelected As Boolean
    Dim pointerKnown As Boolean

    If mKind <> KIND_TRANSFER Then Exit Sub
    If buttonDown Or mRowDragOn Then
        pointerKnown = ReDimUI.FramePointer(pointsX, pointsY)
    End If
    If buttonDown And Not mRowDragPrevDown Then
        mRowDragOn = False
        If mVisible And mEnabled And pointerKnown Then
            If SheetInFront(hostApp) And Not PressClaimedByChrome(hostApp) Then
                If TransferRowOf(PartAtPoint(pointsX, pointsY), panelAt, rowAt) Then
                    mRowDragOn = True
                    mRowDragPanel = panelAt
                    mRowDragFrom = rowAt
                    mRowDragMode = 0
                End If
            End If
        End If
    ElseIf buttonDown And mRowDragOn And pointerKnown Then
        panelAt = TransferPanelAt(pointsX, pointsY)
        If panelAt >= 0 And panelAt <> mRowDragPanel Then
            mRowDragMode = 2
            SetDropPanel panelAt + 1
        Else
            SetDropPanel 0
            If TransferRowOf(PartAtPoint(pointsX, pointsY), panelAt, rowAt) Then
                If panelAt = mRowDragPanel And (rowAt <> mRowDragFrom _
                    Or mRowDragMode = 1) Then
                    mRowDragMode = 1
                    SelectTransferRange mRowDragPanel, mRowDragFrom, rowAt
                End If
            End If
        End If
    ElseIf Not buttonDown And mRowDragOn Then
        mRowDragOn = False
        SetDropPanel 0
        If mRowDragMode = 2 And pointerKnown Then
            If TransferPanelAt(pointsX, pointsY) = 1 - mRowDragPanel Then
                If mRowDragPanel = 0 Then
                    rowSelected = TransferSelectionHas(mSelLeft, mRowDragFrom)
                Else
                    rowSelected = TransferSelectionHas(mSelRight, mRowDragFrom)
                End If
                If rowSelected Then
                    TransferSelected hostApp, (mRowDragPanel = 0)
                Else
                    MoveTransferRow hostApp, (mRowDragPanel = 0), mRowDragFrom
                End If
            End If
        End If
        If mRowDragMode <> 0 Then mHoldSuppressUntilMs = GetTickCount64() + 400
        mRowDragMode = 0
    End If
    mRowDragPrevDown = buttonDown
End Sub

' A row part's panel (0 left, 1 right) and item number.
Private Function TransferRowOf( _
    ByVal partName As String, _
    ByRef panelAt As Long, _
    ByRef rowAt As Long _
) As Boolean
    If Len(partName) < 3 Then Exit Function
    If Not IsNumeric(Mid$(partName, 3)) Then Exit Function
    Select Case Left$(partName, 2)
        Case "al"
            panelAt = 0
            rowAt = TransferShownItem(0, mLeftOffset + CLng(Mid$(partName, 3)))
        Case "cl"
            panelAt = 1
            rowAt = TransferShownItem(1, mRightOffset + CLng(Mid$(partName, 3)))
        Case Else
            Exit Function
    End Select
    TransferRowOf = (rowAt > 0)
End Function

' The panel under a point, rows and empty space alike: 0 left, 1 right,
' -1 neither.
Private Function TransferPanelAt(ByVal pointsX As Double, ByVal pointsY As Double) As Long
    Dim panelW As Double
    Dim panelTop As Double
    Dim panelH As Double
    Dim rightX As Double
    Dim btnX As Double
    Dim btnTop As Double

    TransferPanelAt = -1
    TransferLayout panelW, panelTop, panelH, rightX, btnX, btnTop
    If pointsY < panelTop Or pointsY > panelTop + panelH Then Exit Function
    If pointsX >= mLeft And pointsX <= mLeft + panelW Then
        TransferPanelAt = 0
    ElseIf pointsX >= rightX And pointsX <= rightX + panelW Then
        TransferPanelAt = 1
    End If
End Function

' The panel's selection becomes the rows between two, both included.
Private Sub SelectTransferRange( _
    ByVal panelIndex As Long, _
    ByVal fromRow As Long, _
    ByVal toRow As Long _
)
    Dim rangeRows As Collection
    Dim rowNo As Long
    Dim stepDirection As Long

    Set rangeRows = New Collection
    fromRow = TransferShownPosition(panelIndex, fromRow)
    rowNo = TransferShownPosition(panelIndex, toRow)
    If fromRow = 0 Or rowNo = 0 Then Exit Sub
    stepDirection = IIf(rowNo >= fromRow, 1, -1)
    For rowNo = fromRow To rowNo Step stepDirection
        rangeRows.Add TransferShownItem(panelIndex, rowNo)
    Next rowNo
    If panelIndex = 0 Then
        Set mSelLeft = rangeRows
    Else
        Set mSelRight = rangeRows
    End If
    PlaceKeyCursor panelIndex, toRow
    Touch
End Sub

Private Sub SetDropPanel(ByVal panelNumber As Long)
    If mDropPanel = panelNumber Then Exit Sub
    mDropPanel = panelNumber
    Touch
End Sub

' The drop target panel's edge takes the accent, two points wide.
Private Sub SyncTransferDropTarget( _
    ByVal mainShape As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim panelShape As Shape
    Dim panelNumber As Long

    If mDropPanel = mAppliedDropPanel Then Exit Sub
    On Error Resume Next
    For panelNumber = 1 To 2
        If panelNumber = 1 Then
            Set panelShape = mainShape
        Else
            Set panelShape = hostSheet.Shapes(ShapeName & PART_SEPARATOR & "rp")
        End If
        If panelNumber = mDropPanel Then
            panelShape.Line.ForeColor.RGB = themeValue.PrimaryColor
            panelShape.Line.Weight = 2
        Else
            panelShape.Line.ForeColor.RGB = themeValue.BorderColor
            panelShape.Line.Weight = 1
        End If
    Next panelNumber
    On Error GoTo 0
    mAppliedDropPanel = mDropPanel
End Sub

' A transfer panel's paging arrow moves its window a page less a row.
Private Sub PageTransferPanel(ByVal partName As String)
    Dim transferPage As Long

    EnsureTransferCollections
    transferPage = TransferRowCapacity() - 1
    If transferPage < 1 Then transferPage = 1
    Select Case partName
        Case "alu"
            mLeftOffset = mLeftOffset - transferPage
            If mLeftOffset < 0 Then mLeftOffset = 0
        Case "ald"
            mLeftOffset = mLeftOffset + transferPage
        Case "clu"
            mRightOffset = mRightOffset - transferPage
            If mRightOffset < 0 Then mRightOffset = 0
        Case "cld"
            mRightOffset = mRightOffset + transferPage
        Case Else
            Exit Sub
    End Select
    Touch
End Sub

' Pointer effects set and clear the look; controls that draw none keep
' no state and cost no repaint.
Friend Sub SetPointerLook( _
    ByVal hoveredOn As Boolean, _
    ByVal partName As String, _
    ByVal pressedOn As Boolean _
)
    If hoveredOn Then
        If Not OwnerApp.PointerEffectsOn Then Exit Sub
        Select Case mKind
            Case KIND_BUTTON, KIND_TOGGLE, KIND_TICKBOX, KIND_RADIO, _
                KIND_STEPPER, KIND_SLIDEBAR, KIND_SELECT, KIND_CHECKLIST, _
                KIND_TRANSFER, KIND_TABS, KIND_DATE, KIND_TABLE, KIND_EXPANDER
            Case KIND_INPUT, KIND_COMBO
                If Not IsFloatField() Then Exit Sub
            Case Else
                Exit Sub
        End Select
    ElseIf Not mHovered Then
        Exit Sub
    End If
    If mHovered = hoveredOn And mHoveredPart = partName _
        And mHoverPressed = pressedOn Then Exit Sub
    mHovered = hoveredOn
    mHoveredPart = partName
    mHoverPressed = pressedOn
    Touch
End Sub

Friend Property Get PointerEffectsOn() As Boolean
    PointerEffectsOn = mPointerEffects
End Property

' An open list's highlight follows a moving pointer onto its rows, as a
' Windows list does; Enter from a combo or a keyboard-focused select then
' takes that row.
Friend Sub FollowListPointer(ByVal partName As String)
    Dim rowNumber As Long

    If Not mSelectOpen Then Exit Sub
    If Not OwnerApp.PointerEffectsOn Then Exit Sub
    If Left$(partName, 3) <> "opt" Then Exit Sub
    If Not IsNumeric(Mid$(partName, 4)) Then Exit Sub
    rowNumber = CLng(Mid$(partName, 4))
    If mKind = KIND_SELECT Then
        If Not ItemPickable(ListItemAtRow(rowNumber)) Then Exit Sub
    End If
    If mListOffset + rowNumber = mListHighlight Then Exit Sub
    mListHighlight = mListOffset + rowNumber
    Touch
End Sub

' The pointer look of a fill: under PointerEffects it moves toward its
' ink, 8 percent while the part is hovered and 16 while it is pressed,
' the state layers Material Design draws.
Private Function PointerTint( _
    ByVal baseColor As Long, _
    ByVal inkColor As Long, _
    Optional ByVal partName As String _
) As Long
    PointerTint = baseColor
    If Not mHovered Or mHoveredPart <> partName Then Exit Function
    If Not mEnabled Or mBusy Then Exit Function
    If mHoverPressed Then
        PointerTint = BlendRgb(baseColor, inkColor, 0.16)
    Else
        PointerTint = BlendRgb(baseColor, inkColor, 0.08)
    End If
End Function

' True when the pointer is on the part, for looks that are not tints.
Private Function PointerOn(Optional ByVal partName As String) As Boolean
    PointerOn = mHovered And mHoveredPart = partName And mEnabled And Not mBusy
End Function

' The pointer look as a signature fragment, for the part caches.
Private Function PointerKey() As String
    If mHovered Then PointerKey = mHoveredPart & IIf(mHoverPressed, "!", "?")
End Function

Private Function BlendRgb( _
    ByVal fromColor As Long, _
    ByVal towardColor As Long, _
    ByVal mixWeight As Double _
) As Long
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    redPart = (fromColor And &HFF&) + _
        ((towardColor And &HFF&) - (fromColor And &HFF&)) * mixWeight
    greenPart = ((fromColor \ &H100&) And &HFF&) + _
        (((towardColor \ &H100&) And &HFF&) - ((fromColor \ &H100&) And &HFF&)) * mixWeight
    bluePart = ((fromColor \ &H10000) And &HFF&) + _
        (((towardColor \ &H10000) And &HFF&) - ((fromColor \ &H10000) And &HFF&)) * mixWeight
    BlendRgb = RGB(redPart, greenPart, bluePart)
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
' ink, and a row is brought to the front only when first drawn. A list
' that opens draws its first row and copies it for the rest.
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
    Dim rowIsHeader As Boolean
    Dim boldAt As Long
    Dim boldLength As Long
    Dim priorKey As String
    Dim partCreated As Boolean
    Dim rowCopied As Boolean
    Dim styledRow As Shape
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
        DropPart hostSheet, "optm"
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
        rowIsHeader = IsGroupHeader(itemIndex)
        If rowIsHeader Or Not ItemPickable(itemIndex) Then
            rowInk = themeValue.OnMutedColor
        End If
        rowFrame = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            CStr(rowPitch) & ":" & CStr(resolvedSize)
        If mIsMenu Then
            rowText = CommandGlyphOf(ItemAt(itemIndex)) & vbTab & ItemAt(itemIndex)
        Else
            rowText = GutterText(itemIndex = CLng(mValue), ItemAt(itemIndex))
        End If
        ' A filtering combo bolds what the text matched in each row. The
        ' gutter leads every row and is never bold, so a text write
        ' clears an older match. A group header sits flush left, all
        ' bold.
        boldAt = 0
        boldLength = 0
        If rowIsHeader Then
            rowText = ItemAt(itemIndex)
            If LenB(rowText) > 0 Then
                boldAt = 1
                boldLength = Len(rowText)
            End If
        ElseIf LenB(mListMatchesFilter) > 0 Then
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
            ' A fresh row copies one drawn this pass, then takes only its
            ' own place, fill, words, and ink, as a row whose frame held does.
            Set optionShape = Nothing
            partCreated = False
            If LenB(priorKey) = 0 And Not styledRow Is Nothing Then
                Set optionShape = CopyPartShape(hostSheet, styledRow, "opt" & idx)
            End If
            rowCopied = Not optionShape Is Nothing
            If rowCopied Then
                optionShape.Left = ListSpanLeft()
                optionShape.Top = rowTop
            Else
                Set optionShape = EnsurePartShape( _
                    hostSheet, "opt" & idx, msoShapeRectangle, partCreated)
            End If
            If Not partCreated And Left$(priorKey, Len(rowLook) + 1) _
                = rowLook & Chr$(2) Then
                ' Same row, the match moved: only its letters change.
                MoveRowBold optionShape, Mid$(priorKey, Len(rowLook) + 2), _
                    boldAt, boldLength
            ElseIf rowCopied Or (KeyFrame(priorKey) = rowFrame And Not partCreated) Then
                ' Same frame: only the fill, the text, and the ink move.
                ' A group header's bold first letter would spread over
                ' the text written after it, so grouped lists reset it.
                optionShape.Fill.ForeColor.RGB = rowFill
                With optionShape.TextFrame2.TextRange
                    .Text = rowText
                    .Font.Fill.ForeColor.RGB = rowInk
                    If mIsMenu Then StyleMenuRow optionShape.TextFrame2.TextRange, _
                        themeValue, rowText
                    If Not mGroupHeaders Is Nothing Then .Font.Bold = msoFalse
                    If boldAt > 0 Then
                        .Characters(boldAt, boldLength).Font.Bold = msoTrue
                    End If
                End With
            Else
                optionShape.Left = ListSpanLeft()
                optionShape.Top = rowTop
                optionShape.Width = ListSpanWidth()
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
                If mIsMenu Then StyleMenuRow optionShape.TextFrame2.TextRange, _
                    themeValue, rowText
                If LenB(priorKey) = 0 Or partCreated Then
                    EnsureGutterTab optionShape
                End If
                optionShape.Visible = msoTrue
                optionShape.OnAction = mAppliedOnAction
                If LenB(priorKey) = 0 Then optionShape.ZOrder msoBringToFront
            End If
            Set styledRow = optionShape
        End If
    Next idx
    If wanted = 0 Then
        mListWidth = 0
    ElseIf mIsMenu And mListWidth = 0 Then
        FitMenuWidth hostSheet, wanted
        ' The up pager drew before the rows; it follows the fitted width.
        SyncListPager hostSheet, themeValue, "optu", ChrW(9650), _
            aboveCount, listTopAt, resolvedSize, mAppliedPagerUpSig
    End If
    EnsureRowStates mAppliedRowStates, mAppliedOptionCount
    For idx = wanted + 1 To mAppliedOptionCount
        mAppliedRowStates(idx) = vbNullString
        DropPart hostSheet, "opt" & idx
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
    Dim emptyRowWords As String

    If showIt Then
        If mKind = KIND_COMBO Then
            emptyRowWords = Phrase(UI_NO_MATCHES)
        Else
            emptyRowWords = Phrase(UI_NO_ITEMS)
        End If
        emptyKey = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            CStr(rowPitch) & ":" & CStr(resolvedSize) & ":" & emptyRowWords
        If emptyKey = mAppliedEmptyRowSig Then Exit Sub
        mAppliedEmptyRowSig = emptyKey
        Set emptyShape = EnsurePartShape(hostSheet, "optn", msoShapeRectangle)
        emptyShape.Left = ListSpanLeft()
        emptyShape.Top = rowTop
        emptyShape.Width = ListSpanWidth()
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
            .TextRange.Text = GutterText(False, emptyRowWords)
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
        DropPart hostSheet, "optn"
    End If
End Sub

' Where an open list's rows run across: the face's span, or the width a
' menu fitted to its commands.
Private Function ListSpanLeft() As Double
    If mListWidth > 0 Then
        ListSpanLeft = mListLeft
    Else
        ListSpanLeft = mLeft
    End If
End Function

Private Function ListSpanWidth() As Double
    If mListWidth > 0 Then
        ListSpanWidth = mListWidth
    Else
        ListSpanWidth = mWidth
    End If
End Function

' A menu runs as wide as its widest command needs, never narrower than
' its button, and ends at the button's right edge instead when it would
' pass the window's, as a native menu does. It measures the rows as they
' first draw and gives them the width at once.
Private Sub FitMenuWidth(ByVal hostSheet As Worksheet, ByVal rowTotal As Long)
    Dim idx As Long
    Dim rowNeed As Double
    Dim fitWidth As Double
    Dim fitLeft As Double
    Dim viewLeft As Double
    Dim viewTop As Double
    Dim viewWidth As Double
    Dim viewHeight As Double

    fitWidth = mWidth
    For idx = 1 To rowTotal
        With hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & idx).TextFrame2
            rowNeed = .TextRange.BoundWidth + .MarginLeft + .MarginRight + 6
        End With
        If rowNeed > fitWidth Then fitWidth = rowNeed
    Next idx
    fitLeft = mLeft
    If fitWidth > mWidth Then
        OwnerApp.ResolveViewport viewLeft, viewTop, viewWidth, viewHeight
        If mLeft + fitWidth > viewLeft + viewWidth Then fitLeft = mLeft + mWidth - fitWidth
        If fitLeft < viewLeft Then fitLeft = viewLeft
    End If
    mListLeft = fitLeft
    mListWidth = fitWidth
    If fitWidth = mWidth Then Exit Sub
    For idx = 1 To rowTotal
        With hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & idx)
            .Left = fitLeft
            .Width = fitWidth
        End With
    Next idx
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

    If mKind = KIND_DATE Then
        fullDepth = CAL_DEPTH + 2
    Else
        shownRows = ItemCount
        If shownRows > mListRows Then shownRows = mListRows
        If shownRows < 1 Then shownRows = 1
        fullDepth = shownRows * DropRowPitch()
        If ItemCount > mListRows Then fullDepth = fullDepth + 2 * mHeight
    End If
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
        filterText = FieldTextNow(hostSheet)
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
        currentItem = ExactItemMatch(FieldTextNow(hostSheet))
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
    Dim pagerWords As String

    If beyondCount > 0 Then
        pagerWords = arrowGlyph & " " & Phrase(UI_MORE_ITEMS, beyondCount)
        pagerKey = PartsState(themeValue) & ":" & CStr(rowTop) & ":" & _
            pagerWords & ":" & CStr(resolvedSize) & ":" & CStr(ListSpanLeft()) & ":" & _
            CStr(ListSpanWidth())
        If pagerKey = appliedKey Then Exit Sub
        firstDraw = (LenB(appliedKey) = 0 Or appliedKey = "?")
        appliedKey = pagerKey
        Set pagerShape = EnsurePartShape( _
            hostSheet, partName, msoShapeRectangle)
        pagerShape.Left = ListSpanLeft()
        pagerShape.Top = rowTop
        pagerShape.Width = ListSpanWidth()
        pagerShape.Height = mHeight
        pagerShape.Fill.Visible = msoTrue
        pagerShape.Fill.ForeColor.RGB = themeValue.MutedColor
        pagerShape.Line.Visible = msoTrue
        pagerShape.Line.ForeColor.RGB = themeValue.BorderColor
        pagerShape.Line.Weight = 1
        StyleGlyphText pagerShape, pagerWords, themeValue.OnSurfaceColor, _
            resolvedSize - 1, themeValue
        pagerShape.Visible = msoTrue
        pagerShape.OnAction = mAppliedOnAction
        If firstDraw Then pagerShape.ZOrder msoBringToFront
    ElseIf LenB(appliedKey) > 0 Then
        appliedKey = vbNullString
        DropPart hostSheet, partName
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

Private Function FieldTextNow(ByVal hostSheet As Worksheet) As String
    Dim cell As Range

    If IsFloatField() Then
        FieldTextNow = mText
        Exit Function
    End If
    On Error Resume Next
    Set cell = InputCell(hostSheet)
    If Not cell Is Nothing Then FieldTextNow = CStr(cell.Value)
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
            KIND_TRANSFER, KIND_TABS, KIND_DATE, KIND_TABLE, KIND_EXPANDER
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
    mOffPressPending = False
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
    ' Leaving the command palette cancels it: typed text never runs a
    ' command on the way out, and the palette hides once the blur is done.
    If mComponentId = PALETTE_FIELD_ID Then keepEdits = False
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
                If mKind = KIND_COMBO Then
                    hostApp.SetState mWritesToKey, ItemStateValue(mText)
                Else
                    hostApp.SetState mWritesToKey, mText
                End If
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
    If mComponentId = PALETTE_FIELD_ID Then hostApp.HidePaletteSoon
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
            If HasSelection() And Not mMasked Then
                WriteClipboardText Mid$(mText, SelStart() + 1, SelLength())
            End If
        Case "{CUT}"
            If HasSelection() And Not mMasked Then
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
            currentItem = ExactItemMatch(FieldTextNow(hostApp.Sheet))
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

' The decimal separator Excel types: its own when set apart from the
' system's, else the system's.
Private Function DecimalMark() As String
    If Application.UseSystemSeparators Then
        DecimalMark = Application.International(xlDecimalSeparator)
    Else
        DecimalMark = Application.DecimalSeparator
    End If
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

    separatorChar = DecimalMark()
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
' The clipboard is shared, and another program can hold it open for a
' moment: opening retries for about a tenth of a second before a copy,
' cut, or paste gives up.
Private Function ClaimClipboard() As Boolean
    Dim attemptNo As Long

    For attemptNo = 1 To 10
        If OpenClipboard(0) <> 0 Then
            ClaimClipboard = True
            Exit Function
        End If
        Sleep 10
    Next attemptNo
End Function

Private Function ReadClipboardText() As String
    Dim memoryHandle As LongPtr
    Dim textAddress As LongPtr
    Dim charCount As Long

    If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
    If Not ClaimClipboard() Then Exit Function
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

    If Not ClaimClipboard() Then Exit Function
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
    ByVal pollSelection As Boolean, _
    ByVal buttonDown As Boolean _
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
    If PressedOffUntilRelease(pressEdge, buttonDown) Then LeaveFocus hostApp, True
End Sub

' A press off the control ends its focus or closes its list once the
' button is back up and no click has landed on the control meanwhile.
' Deciding at the press closed a list under a click on its own pager
' whenever the pointer read landed just off the list, and Excel delivers
' that click only on the release. An edge that comes with the button
' already up, as the test seam's does, counts at once.
Private Function PressedOffUntilRelease( _
    ByVal pressEdge As Boolean, _
    ByVal buttonDown As Boolean _
) As Boolean
    If pressEdge Then
        If CursorOverField() Then
            mOffPressPending = False
        ElseIf buttonDown Then
            mOffPressPending = True
            mOffPressUpMs = 0
        Else
            mOffPressPending = False
            PressedOffUntilRelease = True
            Exit Function
        End If
    End If
    If Not mOffPressPending Or buttonDown Then Exit Function
    If mOffPressUpMs = 0 Then mOffPressUpMs = GetTickCount64()
    If GetTickCount64() - mOffPressUpMs < OFF_PRESS_GRACE_MS Then Exit Function
    mOffPressPending = False
    PressedOffUntilRelease = True
End Function

' A click on the control itself: whatever press came before it was on
' the control, however the pointer read.
Friend Sub CancelOffPress()
    mOffPressPending = False
End Sub

' An open list that no field watch covers (a SelectBox, or a cell-backed
' combo's suggestions) closes the way a native drop-down does: on a press
' off its face and rows, or when the grid selection moves. The first frame
' records the selection, since a cell combo opens from an edit whose Enter
' moves the selection right after.
Friend Function NeedsListWatch() As Boolean
    If Not mSelectOpen Or Not mVisible Then Exit Function
    NeedsListWatch = (mKind = KIND_SELECT Or mKind = KIND_DATE) _
        Or (mKind = KIND_COMBO And Not IsFloatField())
End Function

Friend Sub TickListWatch( _
    ByVal pressEdge As Boolean, _
    ByVal pollSelection As Boolean, _
    ByVal buttonDown As Boolean _
)
    If Not NeedsListWatch() Then Exit Sub
    If pollSelection Then
        If LenB(mListSelAtOpen) = 0 Then
            mListSelAtOpen = CurrentSelectionKey()
        ElseIf CurrentSelectionKey() <> mListSelAtOpen Then
            CloseDropList
            Exit Sub
        End If
    End If
    If PressedOffUntilRelease(pressEdge, buttonDown) Then CloseDropList
End Sub

' True when the next apply draws or deletes a whole drop list or calendar:
' its open state changed since it was last drawn.
Friend Property Get DrawsManyParts() As Boolean
    Dim calendarDrawn As Boolean

    Select Case mKind
        Case KIND_SELECT, KIND_COMBO
            DrawsManyParts = (mSelectOpen <> _
                (mAppliedOptionCount > 0 Or LenB(mAppliedEmptyRowSig) > 0))
        Case KIND_DATE
            On Error Resume Next
            calendarDrawn = (LenB(mAppliedRowStates(1)) > 0)
            On Error GoTo 0
            DrawsManyParts = (mSelectOpen <> calendarDrawn)
    End Select
End Property

' Closes an open list without picking anything.
Friend Sub CloseDropList()
    If Not mSelectOpen Then Exit Sub
    mSelectOpen = False
    mListHighlight = 0
    Touch
End Sub

' Every path that opens the list: the app's other lists close, the list
' picks its side, the watch starts over, and the pump runs so the watch
' sees presses elsewhere. A date picker ticks only while its calendar is
' open, so it joins the app's ticking controls here.
Private Sub NoteListOpened()
    mListSelAtOpen = vbNullString
    mOffPressPending = False
    mListWidth = 0
    mListFlipped = ListOpensUpward()
    OwnerApp.CloseListsExcept mComponentId
    If mKind = KIND_DATE Then OwnerApp.TickablesChanged
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
    replacedText = FieldTextNow(hostApp.Sheet)
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
        hostApp.SetState mWritesToKey, ItemStateValue(picked)
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

' Asked of every ticking control each frame, one kind at a time: VBA
' evaluates every term of an Or, so a single expression asked Windows
' about motion and the factory about focus for each control each frame.
Friend Function NeedsAnimation() As Boolean
    Select Case mKind
        Case KIND_SPINNER
            NeedsAnimation = mVisible
        Case KIND_SKELETON
            If mVisible Then NeedsAnimation = Not ReDimUI.MotionReduced
        Case KIND_TOAST
            NeedsAnimation = (mSlideRemaining > 0 Or mToastExiting)
        Case KIND_TOGGLE, KIND_TABS
            NeedsAnimation = (LenB(mGlidePart) > 0)
        Case KIND_INPUT, KIND_COMBO
            If IsFloatField() Then
                NeedsAnimation = ReDimUI.IsComponentFocused(mOwnerAppId, mComponentId)
            End If
    End Select
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
        Case KIND_TOGGLE, KIND_TABS
            If LenB(mGlidePart) > 0 Then TickGlide hostSheet, elapsedMs
        Case KIND_SKELETON
            ' The pulse repaints only when its step changes.
            mPulseMs = mPulseMs + elapsedMs
            If PulseStep() <> mAppliedPulseStep Then Touch
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

Friend Property Get FactoryPointer() As LongPtr
    FactoryPointer = ObjPtr(Me)
End Property

Private Function IsFactory(ByVal instance As ReDimUI) As Boolean
    IsFactory = (ObjPtr(instance) = ReDimUI.FactoryPointer)
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
