Attribute VB_Name = "WhatsNew"
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

' What's New: a tour of what ReDim 1.0.3 adds, a tab for each part, with
' a log on the right that says what every action did.
'
' - Two-way state: WritesTo goes both ways, one listener takes an array
'   of keys, three buttons share a handler through Tag and
'   ReDimUI.Sender.TagValue, and SelectedText and SelectedValue read a
'   pick.
' - Lists: a Filterable select and menu with group headers and disabled
'   items, ItemPosition, lists that scroll, and long items that fit.
' - Forms: Required picks checked by ValidateAll, ErrorText, a DateRange,
'   typed digits, AutoGrow, and a label that follows the field above it.
' - Look: themes from the With builders, text measured in any font,
'   Meiryo and Microsoft YaHei included, the pointer tint, focus rings.
' - Async and errors: a done handler that reads its task's result, two
'   jobs sharing one step, the named error sink, and a Confirm opened
'   from another.
'
' The command palette (Ctrl+Shift+P) turns to any tab.

Private Const APP_ID As String = "whatsnew"
Private Const VIEW_SHEET As String = "WhatsNew"
Private Const LOG_LINES As Long = 12
Private Const SECTIONS_ID As String = "sections"

' Whether the team pick shows the outside error ErrorText sets.
Private gRejectShown As Boolean

Public Sub Auto_Open()
    BuildWhatsNew
End Sub

Public Function WhatsNewApp() As ReDimUI
    Set WhatsNewApp = ReDimUI.App(APP_ID)
End Function

Public Sub BuildWhatsNew()
    Dim ui As ReDimUI
    Dim host As Worksheet

    Set host = ThisWorkbook.Worksheets(1)
    If host.Name <> VIEW_SHEET Then host.Name = VIEW_SHEET
    host.Activate
    Set ui = ReDimUI.Mount(host, APP_ID)
    ' A rebuild over the shapes already drawn applies each control once,
    ' at Render, instead of once per builder call.
    ui.BeginUpdate
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.OnError "WhatsNew.ShowHandlerError"
    ui.CommandPalette
    ui.AddCommand "Show Two-way state", "WhatsNew.HandleShowSection", "View"
    ui.AddCommand "Show Lists", "WhatsNew.HandleShowSection", "View"
    ui.AddCommand "Show Forms", "WhatsNew.HandleShowSection", "View"
    ui.AddCommand "Show Look", "WhatsNew.HandleShowSection", "View"
    ui.AddCommand "Show Async and errors", "WhatsNew.HandleShowSection", "View"

    ui.Label("title").AtRect(24, 14, 420, 30).Text("What's new in ReDim 1.0.3").FontSize(20).Bold
    ui.Label("subtitle").AtRect(24, 44, 620, 18) _
        .Text "Try each tab; the log on the right says what happened. Ctrl+Shift+P jumps to a tab."
    ui.Tabs(SECTIONS_ID).AtRect(24, 70, 620, 30) _
        .Items("Two-way state", "Lists", "Forms", "Look", "Async and errors").WritesTo "section"
    ' A card's words sit at its top, so the log reads down from there.
    ui.Label("logTitle").AtRect(668, 70, 280, 20).Bold.Text "What just happened"
    ui.Card("log").AtRect(668, 94, 280, 446).FontSize(9).BindText "log"

    BuildStateTab ui
    BuildListsTab ui
    BuildFormsTab ui
    BuildLookTab ui
    BuildAsyncTab ui

    ui.SetStateDefault "log", "Nothing yet. Pick, type, or click on the left."
    ui.SetStateDefault "roll", "No roll yet."
    ' One listener on each set of keys.
    ui.OnStateChanged Array("size", "paused", "count"), "WhatsNew.HandleStateChange"
    ui.OnStateChanged Array("themeName", "fontName", "tint"), "WhatsNew.HandleLookChange"
    ui.OnStateChanged Array("hover", "rings"), "WhatsNew.HandlePointerLooks"
    ApplyLookNow ui, False
    ApplyPointerLooksNow ui, False
    ui.Render
    ui.EndUpdate
    ui.ProtectSurface
End Sub

' Two controls writing one key, a switch code turns off, three buttons on
' one handler, and the readers of a pick.
Private Sub BuildStateTab(ByVal ui As ReDimUI)
    ui.Label("stateIntro").AtRect(24, 104, 620, 48).FontSize(10).OnTab(SECTIONS_ID, 1) _
        .Text "WritesTo goes both ways: a control writes its key and follows it. Both " & _
        "controls below write ""size"", so each follows the other, and code that sets the " & _
        "key moves them both."
    ui.SelectBox("sizeSelect").AtRect(24, 172, 160, 24).Caption("SelectBox") _
        .Items("Small", "Medium", "Large").Value(2).WritesTo("size").OnTab SECTIONS_ID, 1
    ui.RadioGroup("sizeRadio").AtRect(210, 172, 140, 66).Caption("RadioGroup") _
        .Items("Small", "Medium", "Large").WritesTo("size").OnTab SECTIONS_ID, 1
    ui.Button("setLarge").AtRect(380, 170, 200, 28).Text("SetState ""size"", ""Large""") _
        .OnClick("WhatsNew.HandleSetLarge").OnTab SECTIONS_ID, 1
    ui.Label("sizeShown").AtRect(380, 204, 260, 18).BindText("size", "The key holds {0}") _
        .OnTab SECTIONS_ID, 1

    ui.Toggle("pause").AtRect(24, 262, 44, 22).Text("Paused").WritesTo("paused") _
        .OnTab SECTIONS_ID, 1
    ui.Button("resume").AtRect(210, 258, 150, 28).Text("Resume from code") _
        .OnClick("WhatsNew.HandleResume").OnTab SECTIONS_ID, 1
    ui.Label("pauseNote").AtRect(380, 258, 260, 42).FontSize(9) _
        .Text("SetState turns the switch off and fires no OnChange.").OnTab SECTIONS_ID, 1

    ui.Label("tagIntro").AtRect(24, 306, 620, 18).Bold _
        .Text("One handler, three buttons: Tag and ReDimUI.Sender.TagValue").OnTab SECTIONS_ID, 1
    ui.Button("add1").AtRect(24, 330, 56, 28).Text("+1").Tag(1) _
        .OnClick("WhatsNew.HandleAdd").OnTab SECTIONS_ID, 1
    ui.Button("add5").AtRect(88, 330, 56, 28).Text("+5").Tag(5) _
        .OnClick("WhatsNew.HandleAdd").OnTab SECTIONS_ID, 1
    ui.Button("add10").AtRect(152, 330, 56, 28).Text("+10").Tag(10) _
        .OnClick("WhatsNew.HandleAdd").OnTab SECTIONS_ID, 1
    ui.Stepper("count").AtRect(230, 330, 130, 28).SliderRange(0, 500, 1).WritesTo("count") _
        .OnTab SECTIONS_ID, 1
    ui.SlideBar("countSlide").AtRect(380, 336, 240, 18).SliderRange(0, 500, 5) _
        .WritesTo("count").OnTab SECTIONS_ID, 1
    ui.Label("countHint").AtRect(24, 366, 620, 18).FontSize(9) _
        .Text("The stepper and the slider follow the count. Focus either one and type " & _
        "digits: 250 sets 250.").OnTab SECTIONS_ID, 1

    ui.Label("readIntro").AtRect(24, 398, 620, 18).Bold _
        .Text("Items with values: SelectedText and SelectedValue").OnTab SECTIONS_ID, 1
    With ui.SelectBox("plan")
        .AtRect(24, 422, 200, 24).Text("Choose a plan").OnTab SECTIONS_ID, 1
        .AddItem "Basic", , 9
        .AddItem "Pro", , 25
        .AddItem "Team", , 60
        .OnChange "WhatsNew.HandlePlan"
    End With
End Sub

' Filtering, group headers, disabled items, finding an item by its text,
' lists that scroll, and long items that fit.
Private Sub BuildListsTab(ByVal ui As ReDimUI)
    ui.Label("listsIntro").AtRect(24, 104, 620, 48).FontSize(10).OnTab(SECTIONS_ID, 2) _
        .Text "Type into the fruit list or the menu to filter it. Durian and Delete are " & _
        "disabled: they show, and no click or key reaches them. Long lists scroll, and a " & _
        "long item fits its list."
    With ui.SelectBox("fruit")
        .AtRect(24, 172, 190, 24).Caption("Filterable SelectBox").Text("Pick a fruit") _
            .OnTab SECTIONS_ID, 2
        .Filterable
        .ListRows 7
        .AddGroup "Everyday"
        .AddItem "Apple"
        .AddItem "Banana"
        .AddItem "Cherry"
        .AddItem "Grape"
        .AddGroup "Exotic"
        .AddItem "Dragon fruit, also sold as pitaya"
        .AddItem "Durian"
        .AddItem "Mangosteen"
        .AddItem "Rambutan"
        .ItemEnabled .ItemPosition("Durian"), False
        .OnChange "WhatsNew.HandleFruit"
    End With
    With ui.MenuButton("actions")
        .AtRect(240, 170, 150, 28).Text("Actions").Filterable.OnTab SECTIONS_ID, 2
        .AddCommand "Copy link", "WhatsNew.HandleMenu", "Link"
        .AddCommand "Export as CSV", "WhatsNew.HandleMenu", "Download"
        .AddCommand "Print", "WhatsNew.HandleMenu", "Print"
        .AddCommand "Share with the team", "WhatsNew.HandleMenu", "Share"
        .AddCommand "Archive", "WhatsNew.HandleMenu", "Folder"
        .AddCommand "Delete", "WhatsNew.HandleMenu", "Delete"
        .ItemEnabled .ItemPosition("Delete"), False
    End With
    ui.Button("findFruit").AtRect(420, 170, 200, 28).Text("Find ""mangosteen""") _
        .OnClick("WhatsNew.HandleFind").OnTab SECTIONS_ID, 2
    ui.ComboBox("city").AtRect(420, 230, 200, 22).Caption("ComboBox; Brussels is disabled") _
        .Items("Amsterdam", "Berlin", "Brussels", "Budapest", "Copenhagen", "Dublin", _
        "Lisbon", "London", "Madrid", "Paris", "Prague", "Rome", "Vienna").OnTab SECTIONS_ID, 2
    ui.ComboBox("city").ItemEnabled ui.ComboBox("city").ItemPosition("Brussels"), False

    ui.RadioGroup("planet").AtRect(24, 230, 180, 110).Caption("A long RadioGroup scrolls") _
        .Items("Mercury", "Venus", "Earth", "Mars", "Jupiter", "Saturn", "Uranus", _
        "Neptune", "Pluto, the dwarf planet", "Ceres", "Eris", "Haumea").OnTab SECTIONS_ID, 2
    ui.CheckList("toppings").AtRect(230, 230, 170, 150).Caption("CheckList: Ctrl+A, or type") _
        .Items("Cheese", "Tomato", "Basil", "Olives", "Mushrooms", "Peppers", "Onion", _
        "Garlic", "Extra mature cheddar, grated", "Pineapple", "Ham", "Spinach", _
        "Anchovies", "Rocket").OnTab SECTIONS_ID, 2
    ui.TransferList("crew").AtRect(24, 400, 380, 120) _
        .Items("Ada", "Grace", "Edsger", "Katherine Johnson of the flight research division", _
        "Margaret", "Alan", "Donald").Captions("Available", "Chosen").Reorderable _
        .OnTab SECTIONS_ID, 2
    ui.Label("listsHint").AtRect(420, 270, 220, 130).FontSize(9) _
        .Text("A long radio group or check list windows its rows behind paging arrows. " & _
        "Page Up, Page Down, and letters move a radio group. A long item ends in an " & _
        "ellipsis inside its control, and an open list runs as wide as its longest " & _
        "item.").OnTab SECTIONS_ID, 2
End Sub

' A form checked by ValidateAll, and a field whose growth moves what sits
' below it.
Private Sub BuildFormsTab(ByVal ui As ReDimUI)
    ui.Label("formsIntro").AtRect(24, 104, 620, 48).FontSize(10).OnTab(SECTIONS_ID, 3) _
        .Text "Submit runs ValidateAll: a Required pick left empty shows its message and " & _
        "takes focus. The date keeps to the next 60 days, and the notes grow with their lines."
    ui.Stack("form").AtRect(24, 172, 290, 0).Gap(6).Stretch.OnTab SECTIONS_ID, 3
    ui.SelectBox("team").Sized(290, 24).Caption("Team").Placeholder("Pick a team") _
        .Items("Design", "Engineering", "Sales", "Support").Required.InStack("form") _
        .OnTab SECTIONS_ID, 3
    ui.DatePicker("start").Sized(290, 24).Caption("Start date").Placeholder("Pick a day") _
        .DateFormat("d mmm yyyy").DateRange(Date, Date + 60).Required.InStack("form") _
        .OnTab SECTIONS_ID, 3
    ui.RadioGroup("contact").Sized(290, 58).Caption("Contact by") _
        .Items("Email", "Phone", "Chat").Required.InStack("form").OnTab SECTIONS_ID, 3
    ui.Toggle("updates").Sized(44, 22).Text("Send me updates").InStack("form") _
        .OnTab SECTIONS_ID, 3
    ui.Stack("formButtons").Across.Gap(8).InStack("form").OnTab SECTIONS_ID, 3
    ui.Button("submit").Sized(100, 30).Text("Submit").Primary _
        .OnClick("WhatsNew.HandleSubmit").InStack("formButtons").OnTab SECTIONS_ID, 3
    ui.Button("reject").Sized(150, 30).Text("Server says no").Secondary _
        .OnClick("WhatsNew.HandleReject").InStack("formButtons").OnTab SECTIONS_ID, 3

    ui.TextInput("notes").AtRect(340, 172, 290, 22).Caption("Notes, with AutoGrow") _
        .Placeholder("Type; Enter starts a new line").AutoGrow(6).OnTab SECTIONS_ID, 3
    ui.Label("followsNotes").Below("notes", 8).Sized(290, 42).FontSize(9) _
        .Text("Placed Below the notes: this label moves down as they grow, and the " & _
        "slider moves with it.").OnTab SECTIONS_ID, 3
    ui.SlideBar("volume").Below("followsNotes", 26).Sized(200, 18) _
        .Caption("SlideBar: focus it and type 75").SliderRange(0, 100, 5).Value(40) _
        .WritesTo("volume").OnTab SECTIONS_ID, 3
    ui.Label("volumeShown").RightOf("volume", 12).Sized(78, 18).BindText("volume", "{0}") _
        .OnTab SECTIONS_ID, 3
End Sub

' The theme, font, and pointer tint the app draws in, and a sample of
' controls to see them on.
Private Sub BuildLookTab(ByVal ui As ReDimUI)
    ui.Label("lookIntro").AtRect(24, 104, 620, 48).FontSize(10).OnTab(SECTIONS_ID, 4) _
        .Text "Themes are built with the With builders. Text is measured in the theme's " & _
        "font, so tabs, cells, badges, and rows fit in any of these fonts, Meiryo and " & _
        "Microsoft YaHei included."
    ui.SelectBox("themeSel").AtRect(24, 172, 180, 24).Caption("Theme") _
        .Items("Light", "Dark", "High contrast", "Ocean", "Midnight").Value(1) _
        .WritesTo("themeName").OnTab SECTIONS_ID, 4
    ui.SelectBox("fontSel").AtRect(230, 172, 180, 24).Caption("Font") _
        .Items("Segoe UI", "Georgia", "Consolas", "Meiryo", "Microsoft YaHei").Value(1) _
        .WritesTo("fontName").OnTab SECTIONS_ID, 4
    ui.SlideBar("tint").AtRect(436, 178, 190, 18).Caption("Hover tint, percent") _
        .SliderRange(0, 30, 2).Value(8).WritesTo("tint").OnTab SECTIONS_ID, 4
    ui.Toggle("hover").AtRect(24, 214, 44, 22).Text("Hover and press looks").Checked(True) _
        .WritesTo("hover").OnTab SECTIONS_ID, 4
    ui.Toggle("rings").AtRect(260, 214, 44, 22).Text("Focus rings").WritesTo("rings") _
        .OnTab SECTIONS_ID, 4

    ui.Label("sampleIntro").AtRect(24, 254, 620, 18).Bold.Text("A sample in the chosen look") _
        .OnTab SECTIONS_ID, 4
    ui.Button("samplePrimary").AtRect(24, 278, 90, 28).Text("Primary").Primary _
        .OnTab SECTIONS_ID, 4
    ui.Button("sampleSecondary").AtRect(122, 278, 96, 28).Text("Secondary").Secondary _
        .OnTab SECTIONS_ID, 4
    ui.Button("sampleSuccess").AtRect(226, 278, 90, 28).Text("Success").Success _
        .OnTab SECTIONS_ID, 4
    ui.Button("sampleWarning").AtRect(324, 278, 90, 28).Text("Warning").Warning _
        .OnTab SECTIONS_ID, 4
    ui.Button("sampleDanger").AtRect(422, 278, 90, 28).Text("Danger").Danger _
        .OnTab SECTIONS_ID, 4
    ui.Badge("sampleBadge").AtRect(526, 283, 0, 18).Text("WIDE NEWS").Warning _
        .OnTab SECTIONS_ID, 4
    ui.Tabs("sampleTabs").AtRect(24, 318, 380, 30) _
        .Items("Overview", "Notifications", "Advanced settings").OnTab SECTIONS_ID, 4
    With ui.Table("sampleTable")
        .AtRect(24, 356, 380, 128).Columns("Item", "Quantity on hand", "Price") _
            .OnTab SECTIONS_ID, 4
        .ColumnWidths 0, 80, 0
        .ColumnFormat 3, "0.00"
        .AddRow "Widget", 120, 2.5
        .AddRow "WWW-XL wide gadget", 8, 14
        .AddRow "Sprocket", 64, 0.75
        .AddRow "Flange", 31, 3.2
        .SortBy 2
    End With
    ui.TextInput("sampleNotes").AtRect(430, 334, 200, 22).Caption("AutoGrow in the font") _
        .AutoGrow(4).OnTab SECTIONS_ID, 4
    ui.TextInput("sampleNotes").InputValue = "One" & vbLf & "Two" & vbLf & "Three"
End Sub

' An op, two jobs, the error sink, and a Confirm that opens another.
Private Sub BuildAsyncTab(ByVal ui As ReDimUI)
    ui.Label("asyncIntro").AtRect(24, 104, 620, 48).FontSize(10).OnTab(SECTIONS_ID, 5) _
        .Text "An op's done handler reads its task's result through ReDimUI.Sender.Task. " & _
        "Two jobs share one step, told apart by their Tag. A failing handler reaches the " & _
        "error sink, which names it."
    ui.Button("roll").AtRect(24, 168, 170, 30).Text("Roll a die in 1 s").Primary _
        .BusyText("Rolling").OnClick("WhatsNew.HandleRoll").OnTab SECTIONS_ID, 5
    ui.Spinner("rollSpin").AtRect(204, 172, 22, 22).Visible(False).OnTab SECTIONS_ID, 5
    ui.Label("rollShown").AtRect(236, 174, 400, 18).BindText("roll", "{0}").OnTab SECTIONS_ID, 5
    ui.Async("dieRoll").Disables("roll").ShowsSpinner("rollSpin").Tag("one six-sided die") _
        .OnDone "WhatsNew.HandleRolled"

    ui.Button("fillA").AtRect(24, 214, 90, 28).Text("Fill A").Tag("jobA") _
        .OnClick("WhatsNew.HandleFill").OnTab SECTIONS_ID, 5
    ui.ProgressBar("barA").AtRect(124, 223, 300, 10).BindValue("jobA").OnTab SECTIONS_ID, 5
    ui.Button("fillB").AtRect(24, 250, 90, 28).Text("Fill B").Tag("jobB") _
        .OnClick("WhatsNew.HandleFill").OnTab SECTIONS_ID, 5
    ui.ProgressBar("barB").AtRect(124, 259, 300, 10).BindValue("jobB").OnTab SECTIONS_ID, 5
    ui.Button("stopJobs").AtRect(440, 232, 120, 28).Text("Cancel both").Secondary _
        .OnClick("WhatsNew.HandleStopJobs").OnTab SECTIONS_ID, 5
    ui.Job("jobA").Steps("WhatsNew.FillStep").PacedMs(50).Tag("jobA") _
        .JobOnDone("WhatsNew.HandleFilled").JobOnCancel "WhatsNew.HandleFillCancelled"
    ui.Job("jobB").Steps("WhatsNew.FillStep").PacedMs(90).Tag("jobB") _
        .JobOnDone("WhatsNew.HandleFilled").JobOnCancel "WhatsNew.HandleFillCancelled"

    ui.Button("oops").AtRect(24, 302, 200, 30).Text("Run a failing handler").Danger _
        .OnClick("WhatsNew.HandleOops").OnTab SECTIONS_ID, 5
    ui.Label("oopsNote").AtRect(236, 300, 400, 42).FontSize(9) _
        .Text("The error goes to the app's OnError sink, which gets the procedure and " & _
        "the control that ran it.").OnTab SECTIONS_ID, 5
    ui.Button("reset").AtRect(24, 346, 200, 30).Text("Reset the demo").Secondary _
        .OnClick("WhatsNew.HandleReset").OnTab SECTIONS_ID, 5
    ui.Label("resetNote").AtRect(236, 346, 400, 42).FontSize(9) _
        .Text("Asks twice: a Confirm opened from another Confirm's OK stays open.") _
        .OnTab SECTIONS_ID, 5
End Sub

' The palette's tab entries share this handler: the entry's words name
' the tab, and the tab strip follows the key it writes.
Public Sub HandleShowSection()
    Dim ui As ReDimUI

    Set ui = ReDimUI.SenderApp
    ui.SetState "section", Mid$(ui.LastCommand, Len("Show ") + 1)
    LogEvent "The palette ran """ & ui.LastCommand & """; the tabs follow ""section"""
End Sub

Public Sub HandleSetLarge()
    WhatsNewApp().SetState "size", "Large"
End Sub

Public Sub HandleResume()
    WhatsNewApp().SetState "paused", False
End Sub

' +1, +5, and +10 share this handler: each button carries its step as
' its Tag.
Public Sub HandleAdd()
    Dim ui As ReDimUI
    Dim countNow As Long

    Set ui = WhatsNewApp()
    LogEvent ReDimUI.SenderId & " clicked, TagValue " & CStr(ReDimUI.Sender.TagValue)
    countNow = CLng(ui.StateOrDefault("count", 0)) + CLng(ReDimUI.Sender.TagValue)
    If countNow > 500 Then countNow = 500
    ui.SetState "count", countNow
End Sub

' One listener on three keys, put on each by OnStateChanged with an array.
Public Sub HandleStateChange()
    Dim ui As ReDimUI

    Set ui = WhatsNewApp()
    LogEvent "State: size " & CStr(ui.StateOrDefault("size", "")) & ", paused " & _
        CStr(ui.StateOrDefault("paused", False)) & ", count " & _
        CStr(ui.StateOrDefault("count", 0))
End Sub

Public Sub HandlePlan()
    With WhatsNewApp().SelectBox("plan")
        LogEvent "SelectedText " & .SelectedText & ", SelectedValue " & CStr(.SelectedValue)
    End With
End Sub

Public Sub HandleFruit()
    LogEvent "Picked " & WhatsNewApp().SelectBox("fruit").SelectedText
End Sub

Public Sub HandleMenu()
    LogEvent "The menu ran """ & ReDimUI.Sender.LastCommand & """"
End Sub

' ItemPosition finds an item by its text, in any case.
Public Sub HandleFind()
    Dim fruitList As ReDimUI
    Dim foundAt As Long

    Set fruitList = WhatsNewApp().SelectBox("fruit")
    foundAt = fruitList.ItemPosition("mangosteen")
    fruitList.Value foundAt
    LogEvent "ItemPosition(""mangosteen"") is " & foundAt & ", counting the group headers"
End Sub

Public Sub HandleSubmit()
    Dim ui As ReDimUI

    Set ui = WhatsNewApp()
    If Not ui.ValidateAll Then
        LogEvent "ValidateAll held the form: each empty pick shows its message"
        Exit Sub
    End If
    LogEvent "Submitted: " & ui.SelectBox("team").SelectedText & ", from " & _
        Format$(ui.DatePicker("start").PickedDate, "d mmm") & ", by " & _
        ui.RadioGroup("contact").SelectedText
    ui.Toast("Submitted.").Success
End Sub

' A message from outside the form, as a server's answer would be:
' ErrorText shows it under the select, and a second press takes it down.
Public Sub HandleReject()
    gRejectShown = Not gRejectShown
    If gRejectShown Then
        WhatsNewApp().SelectBox("team").ErrorText "That team is full this quarter"
        LogEvent "ErrorText shows under the select"
    Else
        WhatsNewApp().SelectBox("team").ErrorText vbNullString
        LogEvent "ErrorText """" takes the message down"
    End If
End Sub

Public Sub HandleLookChange()
    ApplyLookNow WhatsNewApp(), True
End Sub

Public Sub HandlePointerLooks()
    ApplyPointerLooksNow WhatsNewApp(), True
End Sub

' The theme the Look tab asks for: a preset, or one built with the With
' builders, in the chosen font and pointer tint.
Private Sub ApplyLookNow(ByVal ui As ReDimUI, ByVal logIt As Boolean)
    Dim lookTheme As ReDimUI
    Dim themeName As String
    Dim fontFace As String
    Dim tintShare As Double

    themeName = CStr(ui.StateOrDefault("themeName", "Light"))
    fontFace = CStr(ui.StateOrDefault("fontName", "Segoe UI"))
    tintShare = CDbl(ui.StateOrDefault("tint", 8))
    Select Case themeName
        Case "Dark"
            Set lookTheme = ReDimUI.ThemeDark
        Case "High contrast"
            Set lookTheme = ReDimUI.ThemeHighContrast
        Case "Ocean"
            Set lookTheme = ReDimUI.ThemeLight.WithPrimary(RGB(0, 99, 140), RGB(255, 255, 255)) _
                .WithSurface(RGB(248, 252, 254), RGB(16, 42, 60)) _
                .WithMuted(RGB(222, 235, 242), RGB(56, 80, 96)) _
                .WithBorder(RGB(150, 184, 202)).WithCanvas(RGB(233, 243, 248))
        Case "Midnight"
            Set lookTheme = ReDimUI.ThemeDark.WithPrimary(RGB(150, 130, 255), RGB(20, 14, 48)) _
                .WithSurface(RGB(28, 28, 44), RGB(232, 230, 250)) _
                .WithBorder(RGB(78, 74, 110)).WithCanvas(RGB(16, 16, 28))
        Case Else
            Set lookTheme = ReDimUI.ThemeLight
    End Select
    lookTheme.WithFont fontFace, 11
    lookTheme.WithPointerTint tintShare, tintShare * 2
    ui.SetTheme lookTheme
    If logIt Then
        LogEvent "Theme " & themeName & " in " & fontFace & ", pointer tint " & _
            tintShare & " and " & tintShare * 2 & " percent"
    End If
End Sub

Private Sub ApplyPointerLooksNow(ByVal ui As ReDimUI, ByVal logIt As Boolean)
    ui.PointerEffects CBool(ui.StateOrDefault("hover", True))
    ui.FocusRings CBool(ui.StateOrDefault("rings", False))
    If logIt Then
        LogEvent "PointerEffects " & CStr(ui.StateOrDefault("hover", True)) & _
            ", FocusRings " & CStr(ui.StateOrDefault("rings", False))
    End If
End Sub

' The op's task waits a second and then rolls; the done handler reads
' the roll as the task's result, with no module variable holding it.
Public Sub HandleRoll()
    Dim ui As ReDimUI

    Set ui = WhatsNewApp()
    If ui.Async("dieRoll").IsRunning Then Exit Sub
    ui.Async("dieRoll").RunsTask _
        ROneCOne.Task.Delay(1000).ContinueWith(ROneCOne.Func("WhatsNew.RollDie"))
    ui.Async("dieRoll").Start
    LogEvent "Rolling: Async(""dieRoll"").IsRunning is " & CStr(ui.Async("dieRoll").IsRunning)
End Sub

' The roll's continuation, handed the delay's task.
Public Function RollDie(ByVal delayTask As ROneCOne) As Long
    Randomize
    RollDie = Int(Rnd * 6) + 1
End Function

Public Sub HandleRolled()
    WhatsNewApp().SetState "roll", "Rolled a " & CStr(ReDimUI.Sender.Task.Result) & _
        " with " & CStr(ReDimUI.Sender.TagValue) & "."
    LogEvent "The done handler read Sender.Task.Result: " & CStr(ReDimUI.Sender.Task.Result)
End Sub

' Fill A and Fill B share this handler and one step: each button's Tag
' names its job, and each job's Tag names the key its bar reads.
Public Sub HandleFill()
    Dim ui As ReDimUI
    Dim jobId As String

    Set ui = WhatsNewApp()
    jobId = CStr(ReDimUI.Sender.TagValue)
    If ui.Job(jobId).IsRunning Then
        LogEvent jobId & " is already running: Job(""" & jobId & """).IsRunning"
        Exit Sub
    End If
    ui.SetState jobId, 0
    ui.Job(jobId).StartJob
    LogEvent jobId & " started"
End Sub

Public Function FillStep() As Boolean
    Dim ui As ReDimUI
    Dim barKey As String
    Dim filled As Long

    Set ui = WhatsNewApp()
    barKey = CStr(ReDimUI.Sender.TagValue)
    filled = CLng(ui.StateOrDefault(barKey, 0)) + 5
    If filled > 100 Then filled = 100
    ui.SetState barKey, filled
    FillStep = (filled >= 100)
End Function

Public Sub HandleFilled()
    LogEvent ReDimUI.SenderId & " finished; its done handler ran with the job as Sender"
End Sub

Public Sub HandleFillCancelled()
    LogEvent ReDimUI.SenderId & " cancelled"
End Sub

Public Sub HandleStopJobs()
    Dim ui As ReDimUI

    Set ui = WhatsNewApp()
    ui.CancelJob "jobA"
    ui.CancelJob "jobB"
    LogEvent "CancelJob on both; a job at rest takes it as nothing"
End Sub

Public Sub HandleOops()
    LogEvent "HandleOops raises an error"
    Err.Raise 5, "WhatsNew.HandleOops", "The report has no rows to print."
End Sub

' The app's OnError sink: a failed handler's message, naming the
' procedure and the control that ran it.
Public Sub ShowHandlerError(ByVal failureWords As String)
    LogEvent failureWords
    WhatsNewApp().Toast(failureWords).Danger
End Sub

Public Sub HandleReset()
    WhatsNewApp().Confirm "Reset the demo?", _
        "The log, the count, the roll, and the bars go back to the start.", _
        "WhatsNew.HandleResetAgain", vbNullString, "Reset", "Keep"
End Sub

' The first dialog's OK opens a second one, which stays open.
Public Sub HandleResetAgain()
    WhatsNewApp().Confirm "Really reset?", _
        "This dialog opened from the first one's OK.", _
        "WhatsNew.HandleResetDone", vbNullString, "Reset", "Keep"
End Sub

Public Sub HandleResetDone()
    Dim ui As ReDimUI

    Set ui = WhatsNewApp()
    ui.CancelJob "jobA"
    ui.CancelJob "jobB"
    ui.SetState "count", 0
    ui.SetState "jobA", 0
    ui.SetState "jobB", 0
    ui.SetState "roll", "No roll yet."
    ui.SetState "log", "Reset, after two dialogs."
End Sub

' Puts a line at the top of the log and keeps the newest LOG_LINES.
Private Sub LogEvent(ByVal eventWords As String)
    Dim ui As ReDimUI
    Dim logLines() As String
    Dim keptLines As String
    Dim lineNo As Long

    Set ui = WhatsNewApp()
    keptLines = eventWords
    logLines = Split(CStr(ui.StateOrDefault("log", "")), vbLf)
    For lineNo = 0 To UBound(logLines)
        If lineNo >= LOG_LINES - 1 Then Exit For
        keptLines = keptLines & vbLf & logLines(lineNo)
    Next lineNo
    ui.SetState "log", keptLines
End Sub
