Attribute VB_Name = "ExpenseTracker"
' ReDim 1.0.3 (2026-09-26)
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

' Expense Tracker: a small bookkeeping app on one sheet, built from the
' 1.0 controls. A stacked form with captions, a date picker, a category
' select, and a checked amount adds expenses to a hidden data sheet, and
' ValidateAll holds a form with an error back. The table filters as you
' type, copies with Ctrl+C, and deletes a row on double click behind a
' confirm and an Undo toast. A sparkline, a badge, and a budget meter sum
' up the months, and the Actions menu and the command palette
' (Ctrl+Shift+P) run the rest.

Private Const APP_ID As String = "expenses"
Private Const VIEW_SHEET As String = "Expenses"
Private Const DATA_SHEET As String = "ExpenseData"
Private Const EXPORT_SHEET As String = "ExpenseExport"
Private Const MONTHS_SHOWN As Long = 6

' The data row a delete waits on, and the expense a delete took, kept
' for the toast's Undo.
Private gPendingRow As Long
Private gRemovedExpense As Variant

Public Sub Auto_Open()
    BuildExpenseTracker
End Sub

Public Function ExpenseApp() As ReDimUI
    Set ExpenseApp = ReDimUI.App(APP_ID)
End Function

Public Sub BuildExpenseTracker()
    Dim ui As ReDimUI
    Dim host As Worksheet

    Set host = ThisWorkbook.Worksheets(1)
    If host.Name <> VIEW_SHEET Then host.Name = VIEW_SHEET
    host.Activate
    EnsureStore
    Set ui = ReDimUI.Mount(host, APP_ID)
    ' A rebuild over the shapes already drawn applies each control once,
    ' at Render, instead of once per builder call.
    ui.BeginUpdate
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.PointerEffects
    ui.CommandPalette
    ui.AddCommand "Toggle dark mode", "ExpenseTracker.FlipDarkMode", "View"

    ui.Label("title").AtRect(24, 16, 300, 30).Text("Expense Tracker").FontSize(20).Bold
    ui.Label("subtitle").AtRect(24, 46, 440, 16) _
        .Text "Add what you spend; the numbers on the right keep up."
    ui.MenuButton("actions").AtRect(740, 20, 120, 30).Text "Actions"
    With ui.MenuButton("actions")
        .AddCommand "Load sample data", "ExpenseTracker.HandleSample", "Refresh"
        .AddCommand "Export view to a sheet", "ExpenseTracker.HandleExport", "Download"
        .AddCommand "Clear all expenses", "ExpenseTracker.HandleClearAll", "Delete"
    End With

    BuildForm ui
    BuildSummary ui
    BuildViews ui

    ui.SetStateDefault "darkMode", False
    ui.SetStateDefault "budget", 3000
    ui.OnStateChanged "darkMode", "ExpenseTracker.ShowThemeChoice"
    ShowThemeChoice
    RefreshViews ui
    ui.Render
    ui.EndUpdate
    ui.ProtectSurface
End Sub

' The entry form: a card behind a stack of captioned fields, so the
' message under the amount never pushes the note out of place. Enter in
' a field adds the expense.
Private Sub BuildForm(ByVal ui As ReDimUI)
    ui.Card("formCard").AtRect(24, 76, 290, 306).Text "New expense"
    ui.Stack("form").AtRect(40, 110, 258, 0).Gap(6).Stretch
    ui.DatePicker("when").InStack("form").Caption("Date").DateFormat("d mmm yyyy") _
        .PickDate Date
    ui.SelectBox("category").InStack("form").Caption("Category") _
        .Items("Groceries", "Dining", "Transport", "Housing", "Utilities", _
            "Travel", "Fun", "Other").Value 1
    ui.TextInput("amount").InStack("form").Caption("Amount").Placeholder("0.00") _
        .Numeric(allowNegative:=False).Required.Validates "ExpenseTracker.CheckAmount"
    ui.TextInput("note").InStack("form").Caption("Note") _
        .Placeholder("What was it for?").MaxLength 60
    ui.Button("add").InStack("form").Icon("Add").Text("Add expense").Primary _
        .OnClick "ExpenseTracker.HandleAdd"
    ui.DefaultButton "add"

    ui.Expander("budgetBox").AtRect(24, 420, 290, 28).Text "Monthly budget"
    ui.TextInput("budget").AtRect(40, 456, 120, 22).Numeric(allowNegative:=False) _
        .WritesTo("budget").InExpander("budgetBox").OnChange "ExpenseTracker.HandleBudget"
    ui.Toggle("dark").AtRect(24, 500, 44, 22).Text("Dark mode").WritesTo "darkMode"
End Sub

' This month's spend with its change on last month, six months as a
' sparkline, and the budget used.
Private Sub BuildSummary(ByVal ui As ReDimUI)
    ui.Label("monthLbl").AtRect(340, 76, 104, 18).Text "Spent this month"
    ui.Badge("change").AtRect(448, 76, 0, 18).Text "First month"
    ui.Label("monthTotal").AtRect(340, 96, 280, 34).FontSize(22).Bold.BindText "monthTotal"
    ui.Sparkline("trend").AtRect 660, 80, 200, 44
    ui.Label("trendLbl").AtRect(660, 126, 200, 14).FontSize(9).Text "Last six months"
    ui.ProgressBar("budgetMeter").AtRect(340, 150, 520, 10).BindValue "budgetPct"
    ui.Label("budgetLbl").AtRect(340, 164, 520, 16).BindText "budgetText"
End Sub

' The expenses, newest first, and the spending by category, on two tabs.
Private Sub BuildViews(ByVal ui As ReDimUI)
    ui.Tabs("views").AtRect(340, 196, 520, 30).Items "Expenses", "By category"
    ui.Table("list").AtRect(340, 232, 520, 260).Columns("Date", "Category", "Amount", "Note") _
        .EmptyText("No expenses yet. Add one, or load the sample data from Actions.") _
        .OnRowOpen("ExpenseTracker.HandleOpenRow").OnTab "views", 1
    With ui.Table("list")
        .ColumnFormat 1, "d mmm yyyy"
        .ColumnFormat 3, "#,##0.00"
        .SortBy 1, True
    End With
    ui.Label("listHint").AtRect(340, 496, 520, 16).FontSize(9) _
        .Text("Type to filter, Ctrl+C to copy, double-click a row to delete it.") _
        .OnTab "views", 1
    ui.Table("byCategory").AtRect(340, 232, 360, 220).Columns("Category", "Total", "Share") _
        .OnTab "views", 2
    With ui.Table("byCategory")
        .ColumnFormat 2, "#,##0.00"
        .ColumnFormat 3, "0%"
        .SortBy 2, True
    End With
End Sub

' Loads the table from the data sheet and works out the summary: this
' month against last month, six months for the sparkline, the budget
' used, and the totals by category.
Private Sub RefreshViews(ByVal ui As ReDimUI)
    Dim store As Worksheet
    Dim rowTotal As Long
    Dim records As Variant
    Dim monthTotals(1 To MONTHS_SHOWN) As Double
    Dim monthAt As Long
    Dim recordNo As Long
    Dim spent As Double
    Dim kindNames() As String
    Dim kindTotals() As Double
    Dim kindCount As Long
    Dim kindAt As Long
    Dim allTotal As Double
    Dim breakdown() As Variant

    Set store = ExpenseStore()
    rowTotal = StoreRowCount(store)
    ui.BeginUpdate
    ui.Table("list").TableFrom store.Range("A1").Resize(rowTotal + 1, 4)
    If rowTotal > 0 Then
        records = store.Range("A2").Resize(rowTotal, 4).Value
        ReDim kindNames(1 To rowTotal)
        ReDim kindTotals(1 To rowTotal)
        For recordNo = 1 To rowTotal
            spent = CDbl(records(recordNo, 3))
            allTotal = allTotal + spent
            monthAt = MONTHS_SHOWN - MonthsBefore(CDate(records(recordNo, 1)))
            If monthAt >= 1 And monthAt <= MONTHS_SHOWN Then
                monthTotals(monthAt) = monthTotals(monthAt) + spent
            End If
            For kindAt = 1 To kindCount
                If kindNames(kindAt) = CStr(records(recordNo, 2)) Then Exit For
            Next kindAt
            If kindAt > kindCount Then
                kindCount = kindAt
                kindNames(kindAt) = CStr(records(recordNo, 2))
            End If
            kindTotals(kindAt) = kindTotals(kindAt) + spent
        Next recordNo
    End If
    ui.Sparkline("trend").ValuesFrom monthTotals
    ShowMonth ui, monthTotals(MONTHS_SHOWN), monthTotals(MONTHS_SHOWN - 1)

    ReDim breakdown(0 To kindCount, 1 To 3)
    breakdown(0, 1) = "Category"
    breakdown(0, 2) = "Total"
    breakdown(0, 3) = "Share"
    For kindAt = 1 To kindCount
        breakdown(kindAt, 1) = kindNames(kindAt)
        breakdown(kindAt, 2) = kindTotals(kindAt)
        breakdown(kindAt, 3) = kindTotals(kindAt) / allTotal
    Next kindAt
    ui.Table("byCategory").TableFrom breakdown
    ui.EndUpdate
End Sub

' Whole months from the one a date falls in to this one: 0 for a date
' this month, 1 for last month.
Private Function MonthsBefore(ByVal spentOn As Date) As Long
    MonthsBefore = (Year(Date) * 12 + Month(Date)) - (Year(spentOn) * 12 + Month(spentOn))
End Function

' The month's total, the badge that compares it with last month (red
' when spending rose), and the budget meter, red once past the budget.
Private Sub ShowMonth(ByVal ui As ReDimUI, ByVal thisMonth As Double, ByVal lastMonth As Double)
    Dim budgetLimit As Double
    Dim changeShare As Double

    ui.SetState "monthTotal", Format$(thisMonth, "#,##0.00")
    If lastMonth > 0 Then
        changeShare = (thisMonth - lastMonth) / lastMonth
        ui.Badge("change").Text Format$(changeShare, "+0%;-0%;0%") & " on last month"
        If changeShare > 0 Then
            ui.Badge("change").Danger
        Else
            ui.Badge("change").Success
        End If
    Else
        ui.Badge("change").Secondary.Text "First month"
    End If
    budgetLimit = ReadBudget(ui)
    If budgetLimit > 0 Then
        If thisMonth > budgetLimit Then
            ui.SetState "budgetPct", 100
            ui.ProgressBar("budgetMeter").Danger
        Else
            ui.SetState "budgetPct", thisMonth / budgetLimit * 100
            ui.ProgressBar("budgetMeter").Primary
        End If
        ui.SetState "budgetText", Format$(thisMonth / budgetLimit, "0%") & _
            " of the " & Format$(budgetLimit, "#,##0") & " monthly budget"
    Else
        ui.SetState "budgetPct", 0
        ui.SetState "budgetText", "No monthly budget set"
    End If
End Sub

Private Function ReadBudget(ByVal ui As ReDimUI) As Double
    Dim budgetWords As Variant

    budgetWords = ui.StateOrDefault("budget", 0)
    If IsNumeric(budgetWords) Then ReadBudget = CDbl(budgetWords)
End Function

' Adds the form's expense when every field passes, then clears the form
' for the next one.
Public Sub HandleAdd()
    Dim ui As ReDimUI
    Dim spentOn As Variant
    Dim spent As Double
    Dim spendKind As String

    Set ui = ExpenseApp()
    If Not ui.ValidateAll Then Exit Sub
    spentOn = ui.DatePicker("when").PickedDate
    If IsEmpty(spentOn) Then spentOn = Date
    spent = CDbl(ui.TextInput("amount").InputValue)
    spendKind = ui.SelectBox("category").SelectedText
    AppendExpense CDate(spentOn), spendKind, spent, ui.TextInput("note").InputValue
    ui.TextInput("amount").InputValue = vbNullString
    ui.TextInput("note").InputValue = vbNullString
    RefreshViews ui
    ui.Toast("Added " & spendKind & ", " & Format$(spent, "#,##0.00") & ".").Success
    ui.TextInput("amount").Focus
End Sub

' The amount's check: Required reports an empty field, so any text here
' must be a number above zero.
Public Function CheckAmount(ByVal amountText As String) As String
    If Not IsNumeric(amountText) Then
        CheckAmount = "Enter a number such as 12.50"
    ElseIf CDbl(amountText) <= 0 Then
        CheckAmount = "Enter an amount above zero"
    End If
End Function

Public Sub HandleBudget()
    RefreshViews ExpenseApp()
End Sub

Private Sub AppendExpense( _
    ByVal spentOn As Date, _
    ByVal spendKind As String, _
    ByVal spent As Double, _
    ByVal memo As String _
)
    Dim store As Worksheet

    Set store = ExpenseStore()
    store.Cells(StoreRowCount(store) + 2, 1).Resize(1, 4).Value = _
        Array(spentOn, spendKind, spent, memo)
End Sub

' A double click on a row, or Enter on it, asks before deleting it.
Public Sub HandleOpenRow()
    RequestDelete CLng(ReDimUI.Sender.CurrentValue)
End Sub

' Asks to delete a data row: the table numbers its rows in the order the
' data sheet holds them, whatever the sort.
Public Sub RequestDelete(ByVal rowNumber As Long)
    Dim store As Worksheet

    Set store = ExpenseStore()
    If rowNumber < 1 Or rowNumber > StoreRowCount(store) Then Exit Sub
    gPendingRow = rowNumber
    ExpenseApp().Confirm "Delete this expense?", _
        Format$(store.Cells(rowNumber + 1, 1).Value, "d mmm yyyy") & ", " & _
        store.Cells(rowNumber + 1, 2).Value & ", " & _
        Format$(store.Cells(rowNumber + 1, 3).Value, "#,##0.00"), _
        "ExpenseTracker.ConfirmDelete", vbNullString, "Delete", "Keep"
End Sub

Public Sub ConfirmDelete()
    Dim ui As ReDimUI
    Dim store As Worksheet

    If gPendingRow < 1 Then Exit Sub
    Set ui = ExpenseApp()
    Set store = ExpenseStore()
    gRemovedExpense = store.Cells(gPendingRow + 1, 1).Resize(1, 4).Value
    store.Rows(gPendingRow + 1).Delete
    gPendingRow = 0
    RefreshViews ui
    ui.Toast("Expense deleted.").Action "Undo", "ExpenseTracker.UndoDelete"
End Sub

' The toast's Undo: the expense goes back at the end of the data.
Public Sub UndoDelete()
    Dim store As Worksheet

    If IsEmpty(gRemovedExpense) Then Exit Sub
    Set store = ExpenseStore()
    store.Cells(StoreRowCount(store) + 2, 1).Resize(1, 4).Value = gRemovedExpense
    gRemovedExpense = Empty
    RefreshViews ExpenseApp()
    ExpenseApp().Toast "Expense restored."
End Sub

Public Sub HandleSample()
    LoadSampleRows ExpenseStore()
    RefreshViews ExpenseApp()
    ExpenseApp().Toast "Loaded six months of sample expenses."
End Sub

' Writes the rows the table shows, filtered and sorted, to a sheet of
' their own, and stays on the tracker.
Public Sub HandleExport()
    Dim ui As ReDimUI
    Dim exportSheet As Worksheet

    Set ui = ExpenseApp()
    Set exportSheet = EnsureSheet(EXPORT_SHEET)
    exportSheet.Cells.Clear
    ui.Table("list").ExportTo exportSheet.Range("A1")
    exportSheet.Range("A:A").NumberFormat = "d mmm yyyy"
    exportSheet.Range("C:C").NumberFormat = "#,##0.00"
    exportSheet.Columns("A:D").AutoFit
    ui.Toast "Exported " & ui.Table("list").ShownRowCount & " rows to the " & _
        EXPORT_SHEET & " sheet."
End Sub

Public Sub HandleClearAll()
    ExpenseApp().Confirm "Clear every expense?", _
        "This removes all " & StoreRowCount(ExpenseStore()) & " expenses from the workbook.", _
        "ExpenseTracker.ConfirmClearAll", vbNullString, "Clear all", "Keep"
End Sub

Public Sub ConfirmClearAll()
    Dim store As Worksheet
    Dim rowTotal As Long

    Set store = ExpenseStore()
    rowTotal = StoreRowCount(store)
    If rowTotal > 0 Then store.Rows(2).Resize(rowTotal).Delete
    RefreshViews ExpenseApp()
    ExpenseApp().Toast "Cleared."
End Sub

' The palette's command: the switch follows, since it reads the state.
Public Sub FlipDarkMode()
    Dim ui As ReDimUI

    Set ui = ExpenseApp()
    ui.SetState "darkMode", Not CBool(ui.State("darkMode"))
End Sub

Public Sub ShowThemeChoice()
    Dim ui As ReDimUI

    Set ui = ExpenseApp()
    If CBool(ui.State("darkMode")) Then
        ui.SetTheme ReDimUI.ThemeDark
    Else
        ui.SetTheme ReDimUI.ThemeLight
    End If
End Sub

' Six months of made-up expenses, the same every time, none after today.
Private Sub LoadSampleRows(ByVal store As Worksheet)
    Dim kindList As Variant
    Dim memoList As Variant
    Dim priceList As Variant
    Dim records() As Variant
    Dim monthBack As Long
    Dim itemNo As Long
    Dim recordAt As Long
    Dim pick As Long
    Dim spentOn As Date

    kindList = Array("Groceries", "Dining", "Transport", "Utilities", "Fun", _
        "Groceries", "Travel", "Housing")
    memoList = Array("Weekly shop", "Lunch out", "Train pass", "Power bill", "Cinema", _
        "Market", "Weekend away", "Rent")
    priceList = Array(86.4, 23.5, 42, 118.75, 18, 54.2, 240, 950)
    ReDim records(1 To MONTHS_SHOWN * 7, 1 To 4)
    For monthBack = MONTHS_SHOWN - 1 To 0 Step -1
        For itemNo = 1 To 7
            recordAt = recordAt + 1
            ' The rent comes first each month; the rest cycle through the
            ' everyday kinds.
            pick = (monthBack * 3 + itemNo * 5) Mod 7
            If itemNo = 1 Then pick = 7
            spentOn = DateSerial(Year(Date), Month(Date) - monthBack, _
                1 + (itemNo * 4 + monthBack) Mod 27)
            If spentOn > Date Then spentOn = Date
            records(recordAt, 1) = spentOn
            records(recordAt, 2) = kindList(pick)
            records(recordAt, 3) = Round(priceList(pick) * (0.8 + ((itemNo + monthBack) Mod 5) * 0.1), 2)
            records(recordAt, 4) = memoList(pick)
        Next itemNo
    Next monthBack
    store.Cells(StoreRowCount(store) + 2, 1).Resize(UBound(records, 1), 4).Value = records
End Sub

Private Function ExpenseStore() As Worksheet
    Set ExpenseStore = ThisWorkbook.Worksheets(DATA_SHEET)
End Function

' The data sheet, made on first use, hidden, with its header row and the
' sample expenses so a first look has something to show.
Private Sub EnsureStore()
    Dim store As Worksheet

    If SheetExists(DATA_SHEET) Then Exit Sub
    Set store = EnsureSheet(DATA_SHEET)
    store.Range("A1:D1").Value = Array("Date", "Category", "Amount", "Note")
    LoadSampleRows store
    store.Visible = xlSheetHidden
End Sub

' Expenses on the data sheet, under its header row.
Private Function StoreRowCount(ByVal store As Worksheet) As Long
    StoreRowCount = store.Cells(store.Rows.Count, 1).End(xlUp).Row - 1
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim candidate As Worksheet

    For Each candidate In ThisWorkbook.Worksheets
        If candidate.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next candidate
End Function

' The sheet by name, added at the end when missing, leaving the active
' sheet in front.
Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim wasActive As Object

    If SheetExists(sheetName) Then
        Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
        Exit Function
    End If
    Set wasActive = ActiveSheet
    Set EnsureSheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    EnsureSheet.Name = sheetName
    If Not wasActive Is Nothing Then
        wasActive.Activate
    End If
End Function
