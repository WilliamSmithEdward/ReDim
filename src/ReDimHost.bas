Attribute VB_Name = "ReDimHost"
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

' ReDim host module. Excel can only route Shape.OnAction and SetTimer
' callbacks to procedures in a standard module, so this file carries the
' framework's two entry points plus pump lifecycle and crash rails. All real
' behavior lives in ReDimUI.cls.

Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal windowHandle As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr _
) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal windowHandle As LongPtr, _
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
Private gPinCursor As Boolean
Private gCursorPinned As Boolean
' True while the captured keys are bound. Binding is some ninety OnKey
' calls, so focus moving between controls keeps them instead of paying
' a release and a rebind.
Private gKeysBound As Boolean

' Keyboard capture target. Application.OnKey can only call a standard
' module procedure, so focused text entry routes every character through
' here into the runtime. Never raises.
Public Sub RdxKeyChar(ByVal keyText As String)
    On Error Resume Next
    Err.Clear
    ReDimUI.DispatchKey keyText
    If Err.Number <> 0 Then
        Err.Clear
        ReDimUI.NoteTickFault
    End If
End Sub

' The keys a focused control captures: OnKey codes paired with the key
' text RdxKeyChar dispatches. Letters bind twice so Shift yields
' capitals; then the digits and typing punctuation, and the editing and
' navigation keys. Binding and release walk this one list, so they can
' never disagree about what is bound.
Private Function CapturedKeys() As Collection
    Dim keyTable As Collection
    Dim code As Long

    Set keyTable = New Collection
    For code = Asc("a") To Asc("z")
        keyTable.Add Array(Chr$(code), Chr$(code))
        keyTable.Add Array("+" & Chr$(code), UCase$(Chr$(code)))
    Next code
    For code = Asc("0") To Asc("9")
        keyTable.Add Array(Chr$(code), Chr$(code))
    Next code
    ' The numeric keypad, by virtual-key code: without these its digits
    ' and operators went past a focused field to the grid. The decimal
    ' key travels as a name and types the separator Excel uses.
    For code = 0 To 9
        keyTable.Add Array("{" & CStr(96 + code) & "}", CStr(code))
    Next code
    keyTable.Add Array("{106}", "*")
    keyTable.Add Array("{107}", "+")
    keyTable.Add Array("{109}", "-")
    keyTable.Add Array("{110}", "{DECIMAL}")
    keyTable.Add Array("{111}", "/")
    keyTable.Add Array(" ", " ")
    keyTable.Add Array("-", "-")
    keyTable.Add Array(".", ".")
    keyTable.Add Array(",", ",")
    keyTable.Add Array("{BS}", "{BS}")
    keyTable.Add Array("{DEL}", "{DEL}")
    keyTable.Add Array("{LEFT}", "{LEFT}")
    keyTable.Add Array("{RIGHT}", "{RIGHT}")
    keyTable.Add Array("{UP}", "{UP}")
    keyTable.Add Array("{DOWN}", "{DOWN}")
    keyTable.Add Array("{HOME}", "{HOME}")
    keyTable.Add Array("{END}", "{END}")
    keyTable.Add Array("{PGUP}", "{PGUP}")
    keyTable.Add Array("{PGDN}", "{PGDN}")
    keyTable.Add Array("{ENTER}", "{ENTER}")
    keyTable.Add Array("~", "{ENTER}")
    keyTable.Add Array("^{ENTER}", "{CTRLENTER}")
    keyTable.Add Array("^~", "{CTRLENTER}")
    keyTable.Add Array("{TAB}", "{TAB}")
    keyTable.Add Array("+{TAB}", "{BACKTAB}")
    keyTable.Add Array("{ESC}", "{ESC}")
    keyTable.Add Array("%{DOWN}", "{ALTDOWN}")
    keyTable.Add Array("%{UP}", "{ALTUP}")
    keyTable.Add Array("{F4}", "{F4}")
    ' Selection, word moves and deletes, and the clipboard and undo
    ' chords. Shift variants dispatch with a SHIFT prefix.
    keyTable.Add Array("+{LEFT}", "{SHIFTLEFT}")
    keyTable.Add Array("+{RIGHT}", "{SHIFTRIGHT}")
    keyTable.Add Array("+{UP}", "{SHIFTUP}")
    keyTable.Add Array("+{DOWN}", "{SHIFTDOWN}")
    keyTable.Add Array("+{HOME}", "{SHIFTHOME}")
    keyTable.Add Array("+{END}", "{SHIFTEND}")
    keyTable.Add Array("^{LEFT}", "{WORDLEFT}")
    keyTable.Add Array("^{RIGHT}", "{WORDRIGHT}")
    keyTable.Add Array("^+{LEFT}", "{SHIFTWORDLEFT}")
    keyTable.Add Array("^+{RIGHT}", "{SHIFTWORDRIGHT}")
    keyTable.Add Array("^{HOME}", "{TEXTHOME}")
    keyTable.Add Array("^{END}", "{TEXTEND}")
    keyTable.Add Array("^+{HOME}", "{SHIFTTEXTHOME}")
    keyTable.Add Array("^+{END}", "{SHIFTTEXTEND}")
    keyTable.Add Array("^{BS}", "{WORDBS}")
    keyTable.Add Array("^{DEL}", "{WORDDEL}")
    keyTable.Add Array("^a", "{SELECTALL}")
    keyTable.Add Array("^c", "{COPY}")
    keyTable.Add Array("^x", "{CUT}")
    keyTable.Add Array("^v", "{PASTE}")
    keyTable.Add Array("^z", "{UNDO}")
    keyTable.Add Array("^y", "{REDO}")
    keyTable.Add Array("^+z", "{REDO}")
    keyTable.Add Array("+ ", " ")
    ' Punctuation and symbols, bound by character so Excel maps each to
    ' its key on the active layout. OnKey's own metacharacters go in
    ' braces; the apostrophe and the quote travel as names, since they
    ' would break the quoted procedure string.
    keyTable.Add Array("!", "!")
    keyTable.Add Array("@", "@")
    keyTable.Add Array("#", "#")
    keyTable.Add Array("$", "$")
    keyTable.Add Array("{%}", "%")
    keyTable.Add Array("{^}", "^")
    keyTable.Add Array("&", "&")
    keyTable.Add Array("*", "*")
    keyTable.Add Array("{(}", "(")
    keyTable.Add Array("{)}", ")")
    keyTable.Add Array("_", "_")
    keyTable.Add Array("{+}", "+")
    keyTable.Add Array("=", "=")
    keyTable.Add Array("{[}", "[")
    keyTable.Add Array("{]}", "]")
    keyTable.Add Array("{{}", "{")
    keyTable.Add Array("{}}", "}")
    keyTable.Add Array("\", "\")
    keyTable.Add Array("|", "|")
    keyTable.Add Array(";", ";")
    keyTable.Add Array(":", ":")
    keyTable.Add Array("'", "{APOS}")
    keyTable.Add Array("""", "{QUOTE}")
    keyTable.Add Array("<", "<")
    keyTable.Add Array(">", ">")
    keyTable.Add Array("/", "/")
    keyTable.Add Array("?", "?")
    keyTable.Add Array("`", "`")
    keyTable.Add Array("{~}", "~")
    Set CapturedKeys = keyTable
End Function

' Arms capture for a focused control.
Public Sub RdxBindKeys()
    Dim binding As Variant

    If gKeysBound Then Exit Sub
    On Error Resume Next
    For Each binding In CapturedKeys()
        Application.OnKey binding(0), _
            "'RdxKeyChar """ & Replace(binding(1), """", """""") & """'"
    Next binding
    On Error GoTo 0
    gKeysBound = True
End Sub

' Panic release: restores every key ReDim may have bound, whether or not
' any focus state survives. Safe to call at any time.
Public Sub RdxReleaseKeys()
    Dim binding As Variant

    On Error Resume Next
    For Each binding In CapturedKeys()
        Application.OnKey binding(0)
    Next binding
    On Error GoTo 0
    gKeysBound = False
End Sub

' Target for a control's Shortcut, bound only while an app sheet that
' declares it is in front. Never raises.
Public Sub RdxShortcut(ByVal keyCode As String)
    On Error Resume Next
    Err.Clear
    ReDimUI.DispatchShortcut keyCode
    If Err.Number <> 0 Then
        Err.Clear
        ReDimUI.NoteTickFault
    End If
End Sub

' Alt+letter target for access keys, bound only while an app sheet with
' access keys is in front. Never raises.
Public Sub RdxAccessKey(ByVal keyLetter As String)
    On Error Resume Next
    Err.Clear
    ReDimUI.DispatchAccessKey keyLetter
    If Err.Number <> 0 Then
        Err.Clear
        ReDimUI.NoteTickFault
    End If
End Sub

' The command palette's key: opens the palette of the app on the active
' sheet. Never raises.
Public Sub RdxOpenPalette()
    On Error Resume Next
    Err.Clear
    ReDimUI.OpenPaletteOnActiveSheet
    If Err.Number <> 0 Then
        Err.Clear
        ReDimUI.NoteTickFault
    End If
End Sub

' The command palette field's OnChange: runs the entry picked.
Public Sub RdxPalettePick()
    ReDimUI.RunPalettePick
End Sub

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
    ByVal windowHandle As LongPtr, _
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
        RdxApplyCursorPin
    End If
End Sub

Public Sub RdxStopPump()
    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If
    RdxReleaseCursorPin
    If gTimerResolutionRaised Then
        timeEndPeriod 1
        gTimerResolutionRaised = False
    End If
    RdxClearStoredTimerId
End Sub

' Cursor pinning is opt-in. Pinning suppresses Excel's busy-cursor flip
' during each tick, but it also overrides context cursors, hiding the
' hover hand on interactive shapes while work runs. Lean frames made the
' strobe negligible, so hover affordance wins by default.
Public Sub RdxSetCursorPin(ByVal pinOn As Boolean)
    gPinCursor = pinOn
    If gTimerId <> 0 Then
        If pinOn Then
            RdxApplyCursorPin
        Else
            RdxReleaseCursorPin
        End If
    End If
End Sub

Private Sub RdxApplyCursorPin()
    If Not gPinCursor Then Exit Sub
    On Error Resume Next
    Application.Cursor = xlNorthwestArrow
    gCursorPinned = True
    On Error GoTo 0
End Sub

Private Sub RdxReleaseCursorPin()
    If Not gCursorPinned Then Exit Sub
    On Error Resume Next
    Application.Cursor = xlDefault
    On Error GoTo 0
    gCursorPinned = False
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
