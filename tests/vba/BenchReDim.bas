Attribute VB_Name = "BenchReDim"
Option Explicit

' Performance scenarios for tools/bench.py. Each scenario builds its own
' sheet, runs inside one VBA call (the harness rule), and returns
' "metric=milliseconds|..." timed with QueryPerformanceCounter. Per-op
' metrics are averages; the runner prints them as a table.

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef counter As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef ticksPerSecond As LongLong) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Private gTranscript As String

Private Function NowMs() As Double
    Static ticksPerSecond As LongLong
    Dim counter As LongLong

    If ticksPerSecond = 0 Then QueryPerformanceFrequency ticksPerSecond
    QueryPerformanceCounter counter
    NowMs = counter * 1000# / ticksPerSecond
End Function

Private Sub Record(ByVal metric As String, ByVal elapsedMs As Double)
    If LenB(gTranscript) > 0 Then gTranscript = gTranscript & "|"
    gTranscript = gTranscript & metric & "=" & Format$(elapsedMs, "0.000")
End Sub

' Every run starts from a framework with no apps, so a repeat in the same
' session repeats the same work: a reused app keeps its components, and
' their click debounce would swallow the next run's first clicks.
Private Sub StartScenario()
    gTranscript = vbNullString
    ReDimUI.Shutdown
    ReDimUI.AutoPump False
End Sub

Private Function NewCanvas() As Worksheet
    Set NewCanvas = ActiveWorkbook.Worksheets.Add
End Function

Private Function NumberedItems(ByVal total As Long) As Variant
    Dim entries() As Variant
    Dim idx As Long

    ReDim entries(1 To total)
    For idx = 1 To total
        entries(idx) = "Item" & Format$(idx, "0000")
    Next idx
    NumberedItems = entries
End Function

' Build, first render, idle re-render, binding fan-out, a progress tween,
' a theme swap, and teardown over a 100-component surface.
Public Function BenchRender() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double

    StartScenario
    Set host = NewCanvas()
    Set ui = ReDimUI.Mount(host, "benchr")
    started = NowMs()
    For idx = 1 To 40
        ui.Button("b" & idx).AtRect(10 + (idx Mod 8) * 100, 10 + (idx \ 8) * 40, _
            90, 30).Text("Button " & idx)
    Next idx
    For idx = 1 To 30
        ui.Label("l" & idx).AtRect(10 + (idx Mod 6) * 130, 260 + (idx \ 6) * 24, _
            120, 20).BindText "shared", "Value {0}"
    Next idx
    For idx = 1 To 10
        ui.ProgressBar("p" & idx).AtRect(10, 420 + idx * 16, 200, 10).BindValue "pct"
    Next idx
    For idx = 1 To 10
        ui.TickBox("t" & idx).AtRect(230, 420 + idx * 20, 120, 18).Text("Tick " & idx)
    Next idx
    For idx = 1 To 10
        ui.Toggle("g" & idx).AtRect 370, 420 + idx * 26, 44, 22
    Next idx
    ui.SetState "shared", 0
    ui.SetState "pct", 0
    Record "build100", NowMs() - started

    started = NowMs()
    ui.Render
    Record "firstRender100", NowMs() - started
    started = NowMs()
    ui.Render
    Record "idleRender100", NowMs() - started

    started = NowMs()
    For idx = 1 To 20
        ui.SetState "shared", idx
    Next idx
    Record "fanout30PerSet", (NowMs() - started) / 20

    started = NowMs()
    For idx = 1 To 20
        ui.SetState "pct", idx * 5
    Next idx
    Record "progress10PerSet", (NowMs() - started) / 20

    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchr_t1"
    Record "tickClick", NowMs() - started

    started = NowMs()
    ui.SetTheme ReDimUI.ThemeDark
    Record "themeSwap100", NowMs() - started

    started = NowMs()
    ui.Unmount True
    Record "unmount100", NowMs() - started
    ReDimUI.AutoPump True
    BenchRender = gTranscript
End Function

' The typing hot path: keystrokes into a focused float TextInput, a
' filtering combo over 300 items, and a multi-line field.
Public Function BenchTyping() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double

    StartScenario
    Set host = NewCanvas()
    Set ui = ReDimUI.Mount(host, "bencht")
    ui.TextInput("name").AtRect 24, 24, 220, 20
    ui.ComboBox("pick").AtRect 24, 60, 220, 20
    ui.ComboBox("pick").ItemsFrom NumberedItems(300)
    ui.TextInput("notes").AtRect(24, 300, 220, 90).MultiLine True
    ui.Render

    started = NowMs()
    ReDimUI.DispatchShape "rdm_bencht_name"
    Record "focusField", NowMs() - started
    started = NowMs()
    For idx = 1 To 60
        RdxKeyChar Chr$(97 + (idx Mod 26))
    Next idx
    Record "textKey", (NowMs() - started) / 60
    started = NowMs()
    For idx = 1 To 20
        RdxKeyChar "{LEFT}"
    Next idx
    Record "caretMove", (NowMs() - started) / 20
    ' Enter commits and leaves. Tab would move focus into the combo with
    ' its list closed, and the click below would time a caret placement
    ' instead of a focus that opens the list.
    started = NowMs()
    RdxKeyChar "{ENTER}"
    Record "commitBlur", NowMs() - started

    started = NowMs()
    ReDimUI.DispatchShape "rdm_bencht_pick"
    Record "comboFocus300", NowMs() - started
    started = NowMs()
    RdxKeyChar "i"
    RdxKeyChar "t"
    RdxKeyChar "e"
    RdxKeyChar "m"
    RdxKeyChar "0"
    RdxKeyChar "1"
    Record "comboKey300", (NowMs() - started) / 6
    started = NowMs()
    For idx = 1 To 6
        RdxKeyChar "{BS}"
    Next idx
    Record "comboBackspace300", (NowMs() - started) / 6
    started = NowMs()
    For idx = 1 To 20
        RdxKeyChar "{DOWN}"
    Next idx
    Record "comboArrow300", (NowMs() - started) / 20
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"

    ReDimUI.DispatchShape "rdm_bencht_notes"
    started = NowMs()
    For idx = 1 To 40
        If idx Mod 10 = 0 Then
            RdxKeyChar "{ENTER}"
        Else
            RdxKeyChar Chr$(97 + (idx Mod 26))
        End If
    Next idx
    Record "multiLineKey", (NowMs() - started) / 40
    RdxKeyChar "{TAB}"
    RdxReleaseKeys
    ReDimUI.AutoPump True
    BenchTyping = gTranscript
End Function

' Long lists: opening, walking, and paging a 2000-item combo and select,
' and first render, row toggles, paging, and bulk moves in a 1000-item
' transfer list and a 150-item check list.
Public Function BenchLists() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double
    Dim entries As Variant

    StartScenario
    Set host = NewCanvas()
    Set ui = ReDimUI.Mount(host, "benchl")
    entries = NumberedItems(2000)
    ui.ComboBox("pick").AtRect 24, 24, 200, 22
    ui.ComboBox("pick").ItemsFrom entries
    ui.SelectBox("zone").AtRect 260, 24, 160, 24
    ui.SelectBox("zone").ItemsFrom entries
    ui.Render

    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_pick"
    Record "comboOpen2000", NowMs() - started
    started = NowMs()
    For idx = 1 To 30
        RdxKeyChar "{DOWN}"
    Next idx
    Record "comboArrow2000", (NowMs() - started) / 30
    started = NowMs()
    RdxKeyChar "9"
    Record "comboFilterKey2000", NowMs() - started
    RdxKeyChar "{ESC}"
    RdxKeyChar "{ESC}"

    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_zone"
    Record "selectOpen2000", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_zone__optd"
    Record "selectPage2000", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_zone__opt3"
    Record "selectPick2000", NowMs() - started

    ui.TransferList("pool").AtRect 24, 300, 420, 300
    ui.TransferList("pool").ItemsFrom NumberedItems(1000)
    ui.CheckList("rules").AtRect 480, 300, 200, 2400
    ui.CheckList("rules").ItemsFrom NumberedItems(150)
    started = NowMs()
    ui.Render
    Record "renderTransfer1000Check150", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_pool__al2"
    Record "transferRowToggle1000", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_pool__ald"
    Record "transferPage1000", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_pool__mvar"
    Record "transferAllRight1000", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_pool__mval"
    Record "transferAllLeft1000", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_rules__t5"
    Record "checkToggle150", NowMs() - started
    Sleep 200
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchl_rules__mt"
    Record "checkSelectAll150", NowMs() - started
    RdxReleaseKeys
    ReDimUI.AutoPump True
    BenchLists = gTranscript
End Function

' Item mutation APIs on long lists: sources, inserts, removals.
Public Function BenchItems() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double
    Dim picker As ReDimUI

    StartScenario
    Set host = NewCanvas()
    Set ui = ReDimUI.Mount(host, "benchi")
    ui.SelectBox("pick").AtRect 24, 24, 160, 24
    Set picker = ui.Component("pick")
    host.Range("H1:H2000").Formula = "=""Row""&ROW()"
    host.Range("H1:H2000").Value = host.Range("H1:H2000").Value

    started = NowMs()
    picker.ItemsFrom host.Range("H1:H2000")
    Record "itemsFromRange2000", NowMs() - started
    started = NowMs()
    picker.ItemsFrom NumberedItems(2000)
    Record "itemsFromArray2000", NowMs() - started
    ui.Render
    started = NowMs()
    For idx = 1 To 200
        picker.AddItem "Front" & idx, 1
    Next idx
    Record "addItemFront", (NowMs() - started) / 200
    started = NowMs()
    For idx = 1 To 200
        picker.RemoveItem 1
    Next idx
    Record "removeItemFront", (NowMs() - started) / 200
    ReDimUI.AutoPump True
    BenchItems = gTranscript
End Function

' Pump frame and work-tick costs over a 100-component surface: idle, with a
' visible slider (press watch), with a spinner, and with a toast stack.
Public Function BenchPump() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double

    StartScenario
    Set host = NewCanvas()
    host.Activate
    Set ui = ReDimUI.Mount(host, "benchp")
    For idx = 1 To 100
        ui.Button("b" & idx).AtRect(10 + (idx Mod 10) * 80, 10 + (idx \ 10) * 36, _
            70, 28).Text("B" & idx)
    Next idx
    ui.Render

    started = NowMs()
    For idx = 1 To 200
        ReDimUI.TickAll False
    Next idx
    Record "frameIdle100", (NowMs() - started) / 200
    started = NowMs()
    For idx = 1 To 100
        ReDimUI.TickAll True
    Next idx
    Record "workTickIdle100", (NowMs() - started) / 100

    ui.SlideBar("vol").AtRect 10, 420, 200, 18
    ui.Render
    started = NowMs()
    For idx = 1 To 200
        ReDimUI.TickAll False
    Next idx
    Record "frameSlider100", (NowMs() - started) / 200

    ui.Spinner("spin").AtRect(240, 420, 24, 24).Visible True
    started = NowMs()
    For idx = 1 To 200
        ReDimUI.TickAll False
    Next idx
    Record "frameSpinner100", (NowMs() - started) / 200
    ui.Spinner("spin").Visible False

    started = NowMs()
    For idx = 1 To 5
        ui.Toast "Bench toast " & idx, 60000
    Next idx
    Record "toastSpawn", (NowMs() - started) / 5
    started = NowMs()
    For idx = 1 To 60
        ReDimUI.TickAll False
    Next idx
    Record "frameToasts5", (NowMs() - started) / 60
    ui.Unmount True
    ReDimUI.AutoPump True
    BenchPump = gTranscript
End Function

' Window navigation across three windows with nav bars.
Public Function BenchNavigate() As String
    Dim idx As Long
    Dim started As Double
    Dim sheetIndex As Long
    Dim ui As ReDimUI

    StartScenario
    For sheetIndex = 1 To 3
        Set ui = ReDimUI.Mount(NewCanvas(), "benchw" & sheetIndex)
        ui.AsWindow.WindowTitle "Window " & sheetIndex
        ui.NavBar
        For idx = 1 To 20
            ui.Label("l" & idx).AtRect(24, 60 + idx * 22, 200, 20).Text("Row " & idx)
        Next idx
        ui.Render
    Next sheetIndex
    ReDimUI.Navigate "benchw1"
    started = NowMs()
    For idx = 1 To 12
        ReDimUI.Navigate "benchw" & (1 + (idx Mod 3))
    Next idx
    Record "navigate", (NowMs() - started) / 12
    For sheetIndex = 1 To 3
        ReDimUI.App("benchw" & sheetIndex).Sheet.Visible = xlSheetVisible
    Next sheetIndex
    For sheetIndex = 1 To 3
        ReDimUI.App("benchw" & sheetIndex).Unmount True
    Next sheetIndex
    ReDimUI.AutoPump True
    BenchNavigate = gTranscript
End Function

' The 1.0.0 components: a tab switch over panels of ten buttons, a date
' picker's calendar opening, turning a month, and picking a day, a
' skeleton's pulse frames, and a 1000-row table filled from an array,
' sorted by a header, paged, and a row picked.
Public Function BenchComponents() As String
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim idx As Long
    Dim started As Double
    Dim grid() As Variant
    Dim frameTotal As Double

    StartScenario
    Set host = NewCanvas()
    Set ui = ReDimUI.Mount(host, "benchc")
    ui.Tabs("tabs").AtRect(24, 24, 360, 30).Items "One", "Two", "Three"
    For idx = 1 To 30
        ui.Button("b" & idx).AtRect(24 + ((idx - 1) Mod 10) * 40, 70, 36, 24) _
            .Text(CStr(idx)).OnTab "tabs", (idx - 1) \ 10 + 1
    Next idx
    ui.DatePicker("due").AtRect(24, 120, 150, 24).PickDate DateSerial(2026, 9, 22)
    ui.Skeleton("skel").AtRect(420, 120, 200, 60).SkeletonLines 3
    ui.Render

    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_tabs__tb2"
    Record "tabSwitch10", NowMs() - started

    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_due"
    Record "dateOpen", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_due__cn"
    Record "dateMonth", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_due__cd15"
    Record "datePickClose", NowMs() - started

    ' Only the frames count: a Sleep runs a timer tick or two past its ask.
    For idx = 1 To 20
        started = NowMs()
        ReDimUI.PumpOnce
        frameTotal = frameTotal + NowMs() - started
        Sleep 20
    Next idx
    Record "frameSkeleton", frameTotal / 20

    ReDim grid(1 To 1001, 1 To 5)
    grid(1, 1) = "Name"
    grid(1, 2) = "Qty"
    grid(1, 3) = "Price"
    grid(1, 4) = "Date"
    grid(1, 5) = "Zone"
    For idx = 2 To 1001
        grid(idx, 1) = "Item" & Format$(idx - 1, "0000")
        grid(idx, 2) = (idx * 7919) Mod 1000
        grid(idx, 3) = ((idx * 31) Mod 97) / 4
        grid(idx, 4) = DateSerial(2026, 1, 1) + (idx Mod 300)
        grid(idx, 5) = IIf(idx Mod 3 = 0, "North", "South")
    Next idx
    ui.Table("tbl").AtRect 24, 220, 480, 300
    started = NowMs()
    ui.Table("tbl").TableFrom grid
    Record "tableFrom1000", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_tbl__th2"
    Record "tableSort1000", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_tbl__tn"
    Record "tablePage1000", NowMs() - started
    started = NowMs()
    ReDimUI.DispatchShape "rdm_benchc_tbl__tr3"
    Record "tableRowPick1000", NowMs() - started
    ReDimUI.AutoPump True
    BenchComponents = gTranscript
End Function
