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

    app.Label("lblInput").AtRect(24, 280, 200, 16).Text("Cell-backed input").Bold
    app.TextInput("username").At("C20").WritesTo "userName"
    app.Label("inputHint").AtRect(24, 324, 260, 16) _
        .Text("Type in the framed cell and press Enter.")

    app.Label("lblProgress").AtRect(24, 356, 200, 16) _
        .Text("Slider, stepper, and meter share one state key").Bold
    app.SlideBar("volumeslide").AtRect(24, 378, 240, 18) _
        .SliderRange(0, 100, 5).Value(35).WritesTo("volume").BindValue "volume"
    app.ProgressBar("meter").AtRect(24, 404, 240, 12).BindValue "volume"
    app.Label("meterLbl").AtRect(276, 398, 220, 18) _
        .BindText "volume", "Drag, step, or watch: {0}"
    app.Spinner("spin").AtRect 490, 394, 26, 26

    app.Card("inspector").AtRect(24, 436, 560, 110).Text("State inspector")
    app.Label("inspectorBody").AtRect(36, 464, 536, 74).BindText "inspector"

    app.SetState "notifications", False
    app.SetState "region", "North"
    app.SetState "drawnCheck", False
    app.SetState "priority", "Medium"
    app.SetState "volume", 35
    app.SetState "userName", vbNullString
    app.SetState "lastAction", "none yet"
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

    Set app = GalleryApp()
    app.SetState "lastAction", "async work running"
    ignored = ROneCOne.Task.Delay(1200).Await
    app.SetState "lastAction", "async work finished"
End Sub
