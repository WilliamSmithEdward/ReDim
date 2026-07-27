Attribute VB_Name = "WidgetGallery"
Option Explicit

' Widget Gallery: every ReDim component on one sheet, wired to a live state
' inspector so interactions are visible as data, not just pixels.

Private Const APP_ID As String = "gallery"

Public Sub Auto_Open()
    BuildWidgetGallery
End Sub

Public Function GalleryApp() As ReDimUI
    Set GalleryApp = ReDimUI.App(APP_ID)
End Function

Public Sub BuildWidgetGallery()
    Dim app As ReDimUI
    Dim host As Worksheet

    Set host = ThisWorkbook.Worksheets(1)
    Set app = ReDimUI.Mount(host, APP_ID)
    app.ProtectSurface False
    app.PrepareCanvas

    app.Label("title").AtRect(24, 16, 300, 30).Text("Widget Gallery") _
        .FontSize(20).Bold

    app.Label("lblButtons").AtRect(24, 60, 200, 16).Text("Buttons").Bold
    app.Button("primary").AtRect(24, 80, 96, 30).Text("Primary").Primary _
        .OnClick "WidgetGallery.HandlePing"
    app.Button("secondary").AtRect(128, 80, 96, 30).Text("Secondary") _
        .Secondary.OnClick "WidgetGallery.HandlePing"
    app.Button("success").AtRect(232, 80, 96, 30).Text("Success").Success _
        .OnClick "WidgetGallery.HandleToastSuccess"
    app.Button("danger").AtRect(336, 80, 96, 30).Text("Danger").Danger _
        .OnClick "WidgetGallery.HandleModal"
    app.Button("busywork").AtRect(440, 80, 120, 30).Text("Run async work") _
        .OnClickAsync "WidgetGallery.SimulatedWork"
    app.Label("oplog").AtRect(440, 114, 220, 16).BindText "oplog"

    app.Label("lblValues").AtRect(24, 128, 200, 16).Text("Value controls").Bold
    app.Toggle("notify").AtRect(24, 150, 44, 22).WritesTo "notifications"
    app.Label("notifyLbl").AtRect(76, 152, 110, 18).Text("Notifications")
    app.SelectBox("region").AtRect(196, 148, 130, 24) _
        .Items("North", "South", "East", "West").Value(1).WritesTo "region"
    app.SelectBox("region").OnChange "WidgetGallery.HandleRegionChange"

    app.Label("lblDrawn").AtRect(24, 188, 300, 16) _
        .Text("Drawn controls (fully themed)").Bold
    app.TickBox("consent").AtRect(24, 212, 140, 18).Text("Log activity") _
        .WritesTo "drawnCheck"
    app.RadioGroup("priority").AtRect(190, 204, 130, 60) _
        .Items("Low", "Medium", "High").Value(2).WritesTo "priority"
    app.Stepper("volume").AtRect(350, 206, 120, 24).SliderRange(0, 100, 5) _
        .Value(35).WritesTo("volume").BindValue "volume"

    app.Label("lblInput").AtRect(24, 280, 200, 16).Text("Cell-free fields").Bold
    app.TextInput("username").AtRect(24, 304, 150, 22).WritesTo "userName"
    app.ComboBox("fruit").AtRect(196, 304, 150, 22).WritesTo "fruit"
    app.ComboBox("fruit").Items "Apple", "Apricot", "Banana", "Cherry", _
        "Grape", "Grapefruit"
    app.Label("inputHint").AtRect(24, 332, 480, 16) _
        .Text("Click a field and type: the combo filters live, Enter commits, Esc reverts.")

    app.Label("lblTransfer").AtRect(24, 356, 300, 16) _
        .Text("Transfer list (dual listbox)").Bold
    app.Label("lblCheck").AtRect(420, 356, 150, 16) _
        .Text("Checkbox list").Bold
    app.CheckList("options").AtRect 420, 378, 160, 125
    app.CheckList("options").ItemsFrom( _
        Array("Alerts", "Auto-save", "Dark mode", "Sync")) _
        .CheckedFrom(Array("Auto-save")) _
        .WritesTo "options"
    app.TransferList("crew").AtRect 24, 378, 380, 132
    app.TransferList("crew").ItemsFrom( _
        Array("Ada", "Grace", "Edsger", "Alan")) _
        .ChosenFrom(Array("Barbara")) _
        .Captions("Available", "On mission") _
        .WritesTo "crew"
    app.TransferList("crew").OnChange "WidgetGallery.HandleCrewChange"

    app.Label("lblProgress").AtRect(24, 526, 200, 16) _
        .Text("Slider, stepper, and meter share one state key").Bold
    app.SlideBar("volumeslide").AtRect(24, 548, 240, 18) _
        .SliderRange(0, 100, 5).Value(35).WritesTo("volume").BindValue "volume"
    app.ProgressBar("meter").AtRect(24, 574, 240, 12).BindValue "volume"
    app.Label("meterLbl").AtRect(276, 568, 220, 18) _
        .BindText "volume", "Slide, step, or watch: {0}"
    app.Spinner("spin").AtRect 490, 564, 26, 26

    app.Card("inspector").AtRect(24, 606, 560, 110).Text("State inspector")
    app.Label("inspectorBody").AtRect(36, 634, 536, 74).BindText "inspector"

    app.SetState "notifications", False
    app.SetState "region", "North"
    app.SetState "drawnCheck", False
    app.SetState "priority", "Medium"
    app.SetState "volume", 35
    app.SetState "userName", vbNullString
    app.SetState "fruit", vbNullString
    app.SetState "crew", "Barbara"
    app.SetState "options", "Auto-save"
    app.SetState "lastAction", "none yet"
    app.SetState "oplog", "no run yet"
    RefreshInspector
    WireInspector app
    app.Render
    app.ProtectSurface
End Sub

Private Sub WireInspector(ByVal app As ReDimUI)
    app.OnStateChanged "notifications", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "region", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "drawnCheck", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "priority", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "volume", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "userName", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "lastAction", "WidgetGallery.RefreshInspector"
End Sub

Public Sub RefreshInspector()
    Dim app As ReDimUI
    Dim summary As String

    Set app = GalleryApp()
    summary = "notifications = " & CStr(app.State("notifications")) & _
        "   region = " & CStr(app.State("region")) & _
        "   check = " & CStr(app.State("drawnCheck")) & vbLf & _
        "volume = " & CStr(app.State("volume")) & _
        "   priority = " & CStr(app.State("priority")) & _
        "   userName = " & CStr(app.State("userName")) & vbLf & _
        "last action = " & CStr(app.State("lastAction"))
    app.SetState "inspector", summary
End Sub

Public Sub HandleRegionChange()
    GalleryApp().SetState "lastAction", "region picked"
End Sub

Public Sub HandleCrewChange()
    GalleryApp().SetState "lastAction", "crew transferred"
End Sub

Public Sub HandlePing()
    Dim app As ReDimUI

    Set app = GalleryApp()
    app.SetState "lastAction", ReDimUI.SenderId & " clicked"
    app.Toast ReDimUI.SenderId & " clicked.", 1800
End Sub

Public Sub HandleToastSuccess()
    Dim app As ReDimUI

    Set app = GalleryApp()
    app.SetState "lastAction", "success toast"
    app.Toast "Everything saved cleanly.", 2500
End Sub

Public Sub HandleModal()
    GalleryApp().Confirm "Danger zone", _
        "This is the shapes-based modal. No UserForms anywhere.", _
        "WidgetGallery.HandleModalOk", "WidgetGallery.HandleModalCancel"
End Sub

Public Sub HandleModalOk()
    GalleryApp().SetState "lastAction", "modal confirmed"
End Sub

Public Sub HandleModalCancel()
    GalleryApp().SetState "lastAction", "modal canceled"
End Sub

Public Sub SimulatedWork()
    Dim app As ReDimUI
    Dim ignored As Variant
    Dim startedAt As Double

    Set app = GalleryApp()
    startedAt = Timer
    app.SetState "lastAction", "async work running"
    app.SetState "oplog", "op started " & Format$(startedAt, "0.0") & "s"
    ignored = ROneCOne.Task.Delay(1200).Await
    app.SetState "lastAction", "async work finished"
    app.SetState "oplog", "op ran " & Format$(Timer - startedAt, "0.00") & _
        "s (expected 1.2)"
End Sub
