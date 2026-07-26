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

Private Const SHAPE_PREFIX As String = "rdm_"
Private Const PART_SEPARATOR As String = "__"
Private Const ERROR_BASE As Long = vbObjectError + 4200
Private Const ERROR_INVALID_ARGUMENT As Long = ERROR_BASE + 1
Private Const ERROR_INVALID_OPERATION As Long = ERROR_BASE + 2
Private Const ERROR_NOT_FOUND As Long = ERROR_BASE + 3
Private Const CLICK_DEBOUNCE_MS As Long = 150

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
Private Const KIND_CHECKBOX As Long = 7
Private Const KIND_DROPDOWN As Long = 8
Private Const KIND_SLIDER As Long = 9
Private Const KIND_INPUT As Long = 10
Private Const KIND_TOAST As Long = 11
Private Const KIND_OVERLAY As Long = 12
Private Const KIND_SELECT As Long = 13

Private Const VARIANT_PRIMARY As Long = 0
Private Const VARIANT_SECONDARY As Long = 1
Private Const VARIANT_SUCCESS As Long = 2
Private Const VARIANT_DANGER As Long = 3
Private Const VARIANT_NEUTRAL As Long = 4

Private Const REL_NONE As Long = 0
Private Const REL_BELOW As Long = 1
Private Const REL_RIGHT_OF As Long = 2
Private Const PERSIST_PREFIX As String = "rdm_s_"
Private Const REDIM_VERSION As String = "0.2.0"

Private mRole As Long

' Factory fields (meaningful only on the predeclared instance).
Private mApps As Collection
Private mSenderApp As ReDimUI
Private mSenderComponent As ReDimUI
Private mSenderPartName As String
Private mAutoPumpOff As Boolean
Private mInPumpTick As Boolean
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
Private mPersistOn As Boolean
Private mHotKeys As Collection

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
Private mBindTextKey As String
Private mBindTextFormat As String
Private mBindValueKey As String
Private mBindVisibleKey As String
Private mBindEnabledKey As String
Private mWritesToKey As String
Private mLastClickMs As LongLong
Private mExpiresAtMs As LongLong
Private mItems As Collection
Private mItemsStamp As Long
Private mSliderMin As Long
Private mSliderMax As Long
Private mSliderStep As Long
Private mCloseModalOnClick As Boolean
Private mSelectOpen As Boolean
Private mAppliedOptionCount As Long
Private mRelMode As Long
Private mRelTargetId As String
Private mRelGap As Double
Private mGeometryResolving As Boolean

' Applied-property cache. Empty until the first Apply so adopted shapes are
' fully restyled exactly once.
Private mAppliedValid As Boolean
Private mAppliedText As String
Private mAppliedFill As Long
Private mAppliedTextColor As Long
Private mAppliedFontSize As Double
Private mAppliedFontBold As Boolean
Private mAppliedVisible As Boolean
Private mAppliedLeft As Double
Private mAppliedTop As Double
Private mAppliedWidth As Double
Private mAppliedHeight As Double
Private mAppliedOnAction As String
Private mAppliedValue As Double
Private mAppliedChecked As Boolean
Private mAppliedItemsStamp As Long

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

Friend Property Get FactorySenderPart() As String
    FactorySenderPart = mSenderPartName
End Property

' One pump tick across every mounted app. Called by ReDimHost on WM_TIMER and
' by PumpOnce in deterministic tests. Must never raise.
Friend Sub TickAll()
    Dim appValue As ReDimUI

    If Not IsFactory(Me) Then
        ReDimUI.TickAll
        Exit Sub
    End If
    EnsureFactory
    ' Renders inside a tick skip the ScreenUpdating toggle: flipping it at
    ' pump frequency redraws scroll bars and blips the cursor.
    mInPumpTick = True
    On Error Resume Next
    For Each appValue In mApps
        appValue.TickApp GetTickCount64()
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
    ReDimUI.TickAll
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

Public Function Checkbox(ByVal componentId As String) As ReDimUI
    Set Checkbox = GetOrCreateComponent(componentId, KIND_CHECKBOX)
End Function

Public Function Dropdown(ByVal componentId As String) As ReDimUI
    Set Dropdown = GetOrCreateComponent(componentId, KIND_DROPDOWN)
End Function

' Drawn picker: a themed alternative to the native Dropdown, whose face and
' option list use the app font instead of Excel's fixed form-control font.
Public Function SelectBox(ByVal componentId As String) As ReDimUI
    Set SelectBox = GetOrCreateComponent(componentId, KIND_SELECT)
End Function

Public Function Slider(ByVal componentId As String) As ReDimUI
    Set Slider = GetOrCreateComponent(componentId, KIND_SLIDER)
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

    RequireRole ROLE_APP, "FlushDirty"
    If mDirtyIds.Count = 0 Then Exit Sub
    priorScreenUpdating = Application.ScreenUpdating
    If mDirtyIds.Count > 1 And Not ReDimUI.InPumpTick Then
        Application.ScreenUpdating = False
    End If
    On Error GoTo RestoreScreen
    For Each componentId In mDirtyIds
        If TryGetComponent(CStr(componentId), componentValue) Then
            componentValue.Apply mSheet, mTheme, Me
        End If
    Next componentId
    Set mDirtyIds = New Collection
    Application.ScreenUpdating = priorScreenUpdating
    Exit Sub

RestoreScreen:
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim capturedDescription As String
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    Application.ScreenUpdating = priorScreenUpdating
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

Friend Sub TickApp(ByVal nowMs As LongLong)
    Dim opValue As ReDimUI
    Dim jobValue As ReDimUI
    Dim componentValue As ReDimUI
    Dim expiredIds As Collection

    For Each opValue In mAsyncOps
        If opValue.OpIsRunning Then opValue.TickOp Me
    Next opValue
    For Each jobValue In mJobs
        If jobValue.JobIsRunning Then jobValue.TickJob Me, nowMs
    Next jobValue
    Set expiredIds = New Collection
    For Each componentValue In mComponents
        componentValue.TickAnimation mSheet
        If componentValue.IsExpiredToast(nowMs) Then
            expiredIds.Add componentValue.ComponentId
        End If
    Next componentValue
    Dim expiredId As Variant
    For Each expiredId In expiredIds
        If TryGetComponent(CStr(expiredId), componentValue) Then
            componentValue.Remove
        End If
    Next expiredId
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
        Case KIND_CHECKBOX
            mWidth = 110: mHeight = 20
        Case KIND_DROPDOWN
            mWidth = 130: mHeight = 22
        Case KIND_SLIDER
            mWidth = 160: mHeight = 16
            mSliderMin = 0: mSliderMax = 100: mSliderStep = 1
        Case KIND_INPUT
            mWidth = 140: mHeight = 20
        Case KIND_SELECT
            mWidth = 130: mHeight = 24
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

Public Function BindVisible(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindVisible"
    mBindVisibleKey = stateKey
    Touch
    Set BindVisible = Me
End Function

Public Function BindEnabled(ByVal stateKey As String) As ReDimUI
    RequireRole ROLE_COMPONENT, "BindEnabled"
    mBindEnabledKey = stateKey
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
    mItemsStamp = mItemsStamp + 1
    Touch
    Set Items = Me
End Function

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
    InputValue = CStr(InputCell(OwnerApp.Sheet).Value)
End Property

Public Property Let InputValue(ByVal newValue As String)
    Dim priorEvents As Boolean

    RequireRole ROLE_COMPONENT, "InputValue"
    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    InputCell(OwnerApp.Sheet).Value = newValue
    Application.EnableEvents = priorEvents
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

    If mKind <> KIND_INPUT Then Exit Sub
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
    If Not hostApp.ContentBounds(boundLeft, boundTop, boundRight) Then
        hostApp.ResolveViewport boundLeft, boundTop, viewWidth, viewHeight
        boundRight = boundLeft + viewWidth
    End If
    mLeft = boundRight - mWidth
    If mLeft < boundLeft Then mLeft = boundLeft
    mTop = boundTop + (slot - 1) * (mHeight + 6)
    mHasRect = True
    mVisible = True
End Sub

Friend Property Get ToastSlot() As Long
    If mKind = KIND_TOAST Then ToastSlot = CLng(mValue)
End Property

Friend Function IsExpiredToast(ByVal nowMs As LongLong) As Boolean
    If mKind <> KIND_TOAST Then Exit Function
    IsExpiredToast = (nowMs >= mExpiresAtMs)
End Function

Friend Property Get IsPendingToast() As Boolean
    IsPendingToast = (mKind = KIND_TOAST)
End Property

Friend Property Get CountsForBounds() As Boolean
    CountsForBounds = (mKind <> KIND_TOAST And mKind <> KIND_OVERLAY And mVisible)
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
    Dim hostSheet As Worksheet
    Dim hostApp As ReDimUI
    Dim selectedIndex As Long

    Set hostApp = OwnerApp
    Select Case mKind
        Case KIND_TOGGLE
            mChecked = Not mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
            Touch
        Case KIND_CHECKBOX
            Set hostSheet = hostApp.Sheet
            On Error Resume Next
            ' A caption click arrives without the native box having toggled,
            ' so flip it here before reading the value back.
            If ReDimUI.SenderPart = "lbl" Then
                If hostSheet.Shapes(ShapeName).ControlFormat.Value = xlOn Then
                    hostSheet.Shapes(ShapeName).ControlFormat.Value = xlOff
                Else
                    hostSheet.Shapes(ShapeName).ControlFormat.Value = xlOn
                End If
            End If
            mChecked = (hostSheet.Shapes(ShapeName).ControlFormat.Value = xlOn)
            On Error GoTo 0
            mAppliedChecked = mChecked
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mChecked
        Case KIND_DROPDOWN
            Set hostSheet = hostApp.Sheet
            selectedIndex = 0
            On Error Resume Next
            selectedIndex = hostSheet.Shapes(ShapeName).ControlFormat.ListIndex
            On Error GoTo 0
            mValue = selectedIndex
            mAppliedValue = selectedIndex
            If LenB(mWritesToKey) > 0 And selectedIndex >= 1 Then
                hostApp.SetState mWritesToKey, ItemAt(selectedIndex)
            End If
        Case KIND_SLIDER
            Set hostSheet = hostApp.Sheet
            On Error Resume Next
            mValue = hostSheet.Shapes(ShapeName).ControlFormat.Value
            On Error GoTo 0
            mAppliedValue = mValue
            If LenB(mWritesToKey) > 0 Then hostApp.SetState mWritesToKey, mValue
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

Private Function IsFormControlKind() As Boolean
    IsFormControlKind = (mKind = KIND_CHECKBOX Or mKind = KIND_DROPDOWN _
        Or mKind = KIND_SLIDER)
End Function

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
            mVisible = CBool(ownerApp.State(mBindVisibleKey))
        End If
    End If
    If LenB(mBindEnabledKey) > 0 Then
        If ownerApp.HasState(mBindEnabledKey) Then
            mEnabled = CBool(ownerApp.State(mBindEnabledKey))
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
            Set target = hostSheet.Shapes.AddShape( _
                msoShapeArc, mLeft, mTop, mWidth, mHeight)
            On Error Resume Next
            target.Adjustments(1) = 0
            target.Adjustments(2) = 270
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
        Case KIND_CHECKBOX
            Set target = hostSheet.Shapes.AddFormControl( _
                xlCheckBox, CLng(mLeft), CLng(mTop), CLng(mWidth), CLng(mHeight))
        Case KIND_DROPDOWN
            Set target = hostSheet.Shapes.AddFormControl( _
                xlDropDown, CLng(mLeft), CLng(mTop), CLng(mWidth), CLng(mHeight))
        Case KIND_SLIDER
            Set target = hostSheet.Shapes.AddFormControl( _
                xlScrollBar, CLng(mLeft), CLng(mTop), CLng(mWidth), CLng(mHeight))
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

Private Sub ApplyGeometry(ByVal target As Shape)
    If Not mAppliedValid Or mAppliedLeft <> mLeft Then target.Left = mLeft
    If Not mAppliedValid Or mAppliedTop <> mTop Then target.Top = mTop
    If Not mAppliedValid Or mAppliedWidth <> mWidth Then target.Width = mWidth
    If Not mAppliedValid Or mAppliedHeight <> mHeight Then target.Height = mHeight
    mAppliedLeft = mLeft
    mAppliedTop = mTop
    mAppliedWidth = mWidth
    mAppliedHeight = mHeight
    If Not mAppliedValid Or mAppliedVisible <> mVisible Then
        target.Visible = BoolToMso(mVisible)
        mAppliedVisible = mVisible
    End If
End Sub

Private Sub ApplyStyle(ByVal target As Shape, ByVal themeValue As ReDimUI)
    Dim resolvedFill As Long

    If IsFormControlKind() Then Exit Sub
    resolvedFill = ResolveFill(themeValue)
    Select Case mKind
        Case KIND_LABEL
            If Not mAppliedValid Then
                target.Fill.Visible = msoFalse
                target.Line.Visible = msoFalse
            End If
        Case KIND_SPINNER
            If Not mAppliedValid Or mAppliedFill <> resolvedFill Then
                target.Fill.Visible = msoFalse
                target.Line.Visible = msoTrue
                target.Line.ForeColor.RGB = ResolveSpinnerColor(themeValue)
                target.Line.Weight = 3.25
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
        Case KIND_INPUT
            If Not mAppliedValid Then
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
        Case KIND_LABEL, KIND_CARD, KIND_SELECT, KIND_TOAST
            ' Surface-filled kinds read in surface ink regardless of variant;
            ' the variant fall-through below is for filled buttons only.
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
        Or mKind = KIND_INPUT Or IsFormControlKind() Then Exit Sub
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

Private Sub ApplyBehavior(ByVal target As Shape, ByVal hostSheet As Worksheet)
    Dim desiredAction As String

    Select Case mKind
        Case KIND_BUTTON, KIND_TOGGLE, KIND_CHECKBOX, KIND_DROPDOWN, _
            KIND_SLIDER, KIND_TOAST, KIND_OVERLAY, KIND_SELECT
            desiredAction = "'" & hostSheet.Parent.Name & "'!RdxDispatch"
        Case Else
            If LenB(mClickProc) > 0 Or LenB(mAsyncClickProc) > 0 Then
                desiredAction = "'" & hostSheet.Parent.Name & "'!RdxDispatch"
            End If
    End Select
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
        Case KIND_CHECKBOX
            ApplyCheckboxControl target, hostSheet, themeValue
        Case KIND_DROPDOWN
            ApplyDropdownControl target
        Case KIND_SLIDER
            ApplySliderControl target
        Case KIND_INPUT
            ApplyInputCell hostSheet, themeValue
        Case KIND_SELECT
            ApplySelectParts hostSheet, themeValue
    End Select
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

' The native caption stays empty because Excel fixes form-control caption
' fonts at a small size. A themed textbox part carries the visible caption
' and toggles the box when clicked.
Private Sub ApplyCheckboxControl( _
    ByVal target As Shape, _
    ByVal hostSheet As Worksheet, _
    ByVal themeValue As ReDimUI _
)
    Dim caption As Shape
    Dim resolvedSize As Double

    If Not mAppliedValid Then
        target.TextFrame.Characters.Text = vbNullString
    End If
    If Not mAppliedValid Or mAppliedChecked <> mChecked Then
        If mChecked Then
            target.ControlFormat.Value = xlOn
        Else
            target.ControlFormat.Value = xlOff
        End If
        mAppliedChecked = mChecked
    End If
    resolvedSize = mFontSize
    If resolvedSize <= 0 Then resolvedSize = themeValue.BaseFontSize
    Set caption = EnsurePartShape(hostSheet, "lbl", msoShapeRectangle)
    caption.Left = mLeft + 16
    caption.Top = mTop - 1
    caption.Width = mWidth - 16
    caption.Height = mHeight + 2
    caption.Fill.Visible = msoFalse
    caption.Line.Visible = msoFalse
    With caption.TextFrame2
        .MarginLeft = 2
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = mText
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = themeValue.FontName
            .Size = resolvedSize
            .Bold = BoolToMso(mFontBold)
            .Fill.ForeColor.RGB = themeValue.OnSurfaceColor
        End With
    End With
    caption.Visible = BoolToMso(mVisible)
    caption.OnAction = mAppliedOnAction
    mAppliedText = mText
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

Private Sub ApplyDropdownControl(ByVal target As Shape)
    Dim entry As Variant
    Dim desiredIndex As Long

    If Not mItems Is Nothing Then
        If Not mAppliedValid Or mAppliedItemsStamp <> mItemsStamp Then
            target.ControlFormat.RemoveAllItems
            For Each entry In mItems
                target.ControlFormat.AddItem CStr(entry)
            Next entry
            mAppliedItemsStamp = mItemsStamp
            mAppliedValue = 0
        End If
        desiredIndex = CLng(mValue)
        If desiredIndex >= 1 And desiredIndex <= mItems.Count Then
            If Not mAppliedValid Or mAppliedValue <> desiredIndex Then
                target.ControlFormat.ListIndex = desiredIndex
                mAppliedValue = desiredIndex
            End If
        End If
    End If
End Sub

Private Sub ApplySliderControl(ByVal target As Shape)
    Dim clamped As Long

    If Not mAppliedValid Then
        target.ControlFormat.Min = mSliderMin
        target.ControlFormat.Max = mSliderMax
        target.ControlFormat.SmallChange = mSliderStep
        target.ControlFormat.LargeChange = mSliderStep * 10
    End If
    clamped = CLng(mValue)
    If clamped < mSliderMin Then clamped = mSliderMin
    If clamped > mSliderMax Then clamped = mSliderMax
    If Not mAppliedValid Or mAppliedValue <> clamped Then
        target.ControlFormat.Value = clamped
        mAppliedValue = clamped
    End If
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
    End If
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
    For position = 1 To optionCount
        hostSheet.Shapes(ShapeName & PART_SEPARATOR & "opt" & position).Delete
    Next position
    On Error GoTo 0
End Sub

Friend Function NeedsAnimation() As Boolean
    NeedsAnimation = (mKind = KIND_SPINNER And mVisible)
End Function

' Animation bypasses the diff cache: rotation is transient visual state, not
' a declared prop.
Friend Sub TickAnimation(ByVal hostSheet As Worksheet)
    If Not NeedsAnimation() Then Exit Sub
    On Error Resume Next
    hostSheet.Shapes(ShapeName).IncrementRotation 30
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

Friend Sub TickJob(ByVal ownerApp As ReDimUI, ByVal nowMs As LongLong)
    Dim startedAt As LongLong
    Dim stepDone As Variant

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
    startedAt = GetTickCount64()
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
    Loop While GetTickCount64() - startedAt < CLngLng(mJobBudgetMs)
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
