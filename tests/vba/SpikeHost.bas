Attribute VB_Name = "SpikeHost"
Option Explicit

' Architecture spike for ReDim. Proves three load-bearing assumptions under a
' live Excel session before the framework is built:
'   1. A SetTimer callback fires while Excel idles between automation calls.
'   2. ROneCOne's Friend AdvanceTask steps a task to completion from the pump.
'   3. Worksheet shapes can be created, styled, rotated, and wired to OnAction.

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

Private gTimerId As LongPtr
Private gTickCount As Long
Private gPumpErrors As Long
Private gInTick As Boolean
Private gTask As ROneCOne
Private gTaskDoneTick As Long
Private gLastClicked As String

Public Sub PumpCallback( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long _
)
    If gInTick Then Exit Sub
    gInTick = True
    On Error Resume Next
    gTickCount = gTickCount + 1
    If Not gTask Is Nothing Then
        gTask.AdvanceTask
        If gTask.IsCompleted And gTaskDoneTick = 0 Then
            gTaskDoneTick = gTickCount
        End If
    End If
    If Err.Number <> 0 Then gPumpErrors = gPumpErrors + 1
    On Error GoTo 0
    gInTick = False
End Sub

Public Function StartPump(ByVal intervalMs As Long) As Boolean
    If gTimerId <> 0 Then StopPump
    gTickCount = 0
    gPumpErrors = 0
    gTimerId = SetTimer(0, 0, intervalMs, AddressOf PumpCallback)
    StartPump = (gTimerId <> 0)
End Function

Public Function StopPump() As Boolean
    If gTimerId <> 0 Then
        StopPump = (KillTimer(0, gTimerId) <> 0)
        gTimerId = 0
    End If
End Function

Public Function GetTicks() As Long
    GetTicks = gTickCount
End Function

Public Function GetPumpErrors() As Long
    GetPumpErrors = gPumpErrors
End Function

Public Function StartDelayTask(ByVal delayMs As Long) As String
    Set gTask = ROneCOne.Task.Delay(delayMs)
    gTaskDoneTick = 0
    StartDelayTask = gTask.Status
End Function

Public Function GetTaskStatus() As String
    If gTask Is Nothing Then
        GetTaskStatus = "NoTask"
    Else
        GetTaskStatus = gTask.Status
    End If
End Function

Public Function GetTaskDoneTick() As Long
    GetTaskDoneTick = gTaskDoneTick
End Function

Public Function SetGlobal(ByVal value As Long) As Long
    gTickCount = value
    SetGlobal = gTickCount
End Function

Public Function GetGlobal() As Long
    GetGlobal = gTickCount
End Function

' Proves the timer fires while VBA pumps messages, all inside one call so the
' TIMERPROC address cannot go stale between automation round trips.
Public Function PumpProofSingleCall() As String
    Dim ignored As Variant
    Dim ticks As Long

    StartPump 50
    ignored = ROneCOne.Task.Delay(600).Await
    ticks = gTickCount
    StopPump
    PumpProofSingleCall = ticks & "|" & gPumpErrors
End Function

' Proves the pump advances a registered task with no Await on that task.
Public Function TaskPumpProofSingleCall() As String
    Dim ignored As Variant

    Set gTask = ROneCOne.Task.Delay(300)
    gTaskDoneTick = 0
    StartPump 50
    ignored = ROneCOne.Task.Delay(900).Await
    StopPump
    TaskPumpProofSingleCall = gTask.Status & "|" & gTaskDoneTick & "|" & _
        gTickCount & "|" & gPumpErrors
    Set gTask = Nothing
End Function

Public Function CreateSpikeShape() As String
    Dim host As Worksheet
    Dim target As Shape

    Set host = ActiveWorkbook.Worksheets(1)
    On Error Resume Next
    host.Shapes("rdm_spike_button").Delete
    On Error GoTo 0
    Set target = host.Shapes.AddShape(msoShapeRoundedRectangle, 20, 20, 120, 36)
    target.Name = "rdm_spike_button"
    target.TextFrame2.TextRange.Text = "Spike"
    target.Fill.ForeColor.RGB = RGB(33, 115, 70)
    target.Line.Visible = msoFalse
    target.OnAction = "SpikeHost.HandleSpikeClick"
    CreateSpikeShape = target.Name
End Function

Public Function ReadSpikeShape() As String
    Dim target As Shape

    Set target = ActiveWorkbook.Worksheets(1).Shapes("rdm_spike_button")
    ReadSpikeShape = target.Name & "|" & _
        target.TextFrame2.TextRange.Text & "|" & _
        target.OnAction
End Function

Public Function RotateSpikeShape() As Double
    Dim target As Shape

    Set target = ActiveWorkbook.Worksheets(1).Shapes("rdm_spike_button")
    target.IncrementRotation 30
    RotateSpikeShape = target.Rotation
End Function

Public Sub HandleSpikeClick()
    gLastClicked = "rdm_spike_button"
End Sub

Public Function DispatchByName(ByVal shapeName As String) As String
    Dim target As Shape

    Set target = ActiveWorkbook.Worksheets(1).Shapes(shapeName)
    Application.Run target.OnAction
    DispatchByName = gLastClicked
End Function
