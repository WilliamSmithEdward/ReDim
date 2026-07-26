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

' SetTimer rides the global timer resolution, which defaults to roughly
' 15.6 ms and quantizes frames into uneven buckets. Raising it to 1 ms
' while the pump is armed makes frames land on schedule; it is restored on
' stop so the power cost exists only while something animates.
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long _
) As Long

Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" ( _
    ByVal uPeriod As Long _
) As Long

' Animation frame interval. Work (ops, budget jobs, toast expiry) runs on a
' 50 ms cadence inside ReDimUI.TickAll regardless of the frame rate.
Private Const PUMP_DEFAULT_INTERVAL_MS As Long = 16
Private Const PUMP_MAX_CONSECUTIVE_ERRORS As Long = 10
Private Const PUMP_ID_NAME As String = "rdm_pump_id"

Private gTimerId As LongPtr
Private gInTick As Boolean
Private gConsecutiveErrors As Long
Private gTickCount As LongLong
Private gTimerResolutionRaised As Boolean

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
    If gTimerId <> 0 Then
        RdxStoreTimerId gTimerId
        If Not gTimerResolutionRaised Then
            gTimerResolutionRaised = (timeBeginPeriod(1) = 0)
        End If
        ' Excel flips to the busy cursor whenever VBA executes, which at
        ' pump frequency reads as a strobe. Pinning the cursor while the
        ' timer is armed keeps it steady; StopPump restores the default.
        On Error Resume Next
        Application.Cursor = xlNorthwestArrow
        On Error GoTo 0
    End If
End Sub

Public Sub RdxStopPump()
    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
        On Error Resume Next
        Application.Cursor = xlDefault
        On Error GoTo 0
    End If
    If gTimerResolutionRaised Then
        timeEndPeriod 1
        gTimerResolutionRaised = False
    End If
    RdxClearStoredTimerId
End Sub

Public Function RdxPumpArmed() As Boolean
    RdxPumpArmed = (gTimerId <> 0)
End Function

Public Function RdxTickCount() As LongLong
    RdxTickCount = gTickCount
End Function

' Deterministic single tick for tests and debugging: one nominal 50 ms
' frame with the work pass forced, without arming a timer.
Public Sub RdxPumpOnce()
    ReDimUI.PumpOnce
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
