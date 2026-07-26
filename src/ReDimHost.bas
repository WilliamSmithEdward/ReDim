Attribute VB_Name = "ReDimHost"
Option Explicit

' ReDim host module. Excel can only route Shape.OnAction and SetTimer
' callbacks to procedures in a standard module, so this file carries the
' framework's two entry points plus pump lifecycle and crash rails. All real
' behavior lives in ReDimUI.cls.

Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr _
) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr _
) As Long

Private Const PUMP_DEFAULT_INTERVAL_MS As Long = 50
Private Const PUMP_MAX_CONSECUTIVE_ERRORS As Long = 10
Private Const PUMP_ID_NAME As String = "rdm_pump_id"

Private gTimerId As LongPtr
Private gInTick As Boolean
Private gConsecutiveErrors As Long
Private gTickCount As LongLong

' Shape.OnAction target for every ReDim component. Application.Caller carries
' the clicked shape's name.
Public Sub RdxDispatch()
    Dim callerName As Variant

    On Error GoTo SwallowError
    callerName = Application.Caller
    If VarType(callerName) = vbString Then
        ReDimUI.DispatchShape CStr(callerName)
    End If
    Exit Sub

SwallowError:
    ' A dispatch failure must never surface Excel's runtime error dialog.
End Sub

' SetTimer callback. Keep this minimal: one guarded call into the runtime.
' An error escaping a TIMERPROC can take down the Excel process.
Public Sub RdxPumpCallback( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long _
)
    If gInTick Then Exit Sub
    gInTick = True
    On Error Resume Next
    gTickCount = gTickCount + 1
    ReDimUI.TickAll
    If Err.Number <> 0 Then
        gConsecutiveErrors = gConsecutiveErrors + 1
        If gConsecutiveErrors >= PUMP_MAX_CONSECUTIVE_ERRORS Then RdxStopPump
    Else
        gConsecutiveErrors = 0
        If Not ReDimUI.HasPendingWork Then RdxStopPump
    End If
    On Error GoTo 0
    gInTick = False
End Sub

Public Sub RdxEnsurePump(Optional ByVal intervalMs As Long = PUMP_DEFAULT_INTERVAL_MS)
    If gTimerId <> 0 Then Exit Sub
    RdxKillOrphanTimer
    gConsecutiveErrors = 0
    gTimerId = SetTimer(0, 0, intervalMs, AddressOf RdxPumpCallback)
    If gTimerId <> 0 Then RdxStoreTimerId gTimerId
End Sub

Public Sub RdxStopPump()
    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If
    RdxClearStoredTimerId
End Sub

Public Function RdxPumpArmed() As Boolean
    RdxPumpArmed = (gTimerId <> 0)
End Function

Public Function RdxTickCount() As LongLong
    RdxTickCount = gTickCount
End Function

' Deterministic single tick for tests and debugging: same body the timer
' runs, without arming a timer.
Public Sub RdxPumpOnce()
    RdxPumpCallback 0, 0, 0, 0
End Sub

' The armed timer id survives VBA state loss inside a workbook-scoped name,
' so a rebuilt session can kill the orphan before arming a fresh timer.
Private Sub RdxStoreTimerId(ByVal timerId As LongPtr)
    On Error Resume Next
    ThisWorkbook.Names(PUMP_ID_NAME).Delete
    ThisWorkbook.Names.Add PUMP_ID_NAME, "=" & CStr(timerId), False
    On Error GoTo 0
End Sub

Private Sub RdxClearStoredTimerId()
    On Error Resume Next
    ThisWorkbook.Names(PUMP_ID_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub RdxKillOrphanTimer()
    Dim stored As String
    Dim orphanId As LongPtr

    On Error Resume Next
    stored = ThisWorkbook.Names(PUMP_ID_NAME).RefersTo
    On Error GoTo 0
    If LenB(stored) = 0 Then Exit Sub
    stored = Replace(stored, "=", vbNullString)
    If IsNumeric(stored) Then
        orphanId = CLngLng(stored)
        If orphanId <> 0 And orphanId <> gTimerId Then KillTimer 0, orphanId
    End If
    RdxClearStoredTimerId
End Sub

' Best-effort rail: kill the pump when the hosting workbook closes so no
' TIMERPROC outlives its VBA project.
Public Sub Auto_Close()
    RdxStopPump
    ReDimUI.Shutdown
End Sub
