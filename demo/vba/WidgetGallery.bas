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
    app.Checkbox("audit").AtRect(196, 150, 130, 20).Text("Audit trail") _
        .WritesTo "audit"
    app.SelectBox("region").AtRect(340, 148, 130, 24) _
        .Items("North", "South", "East", "West").Value(1).WritesTo "region"
    app.SelectBox("region").OnChange "WidgetGallery.HandleRegionChange"

    app.Label("lblNative").AtRect(24, 188, 300, 16) _
        .Text("Native controls (Excel-drawn, fixed font)").Bold
    app.Dropdown("nativedd").AtRect(24, 210, 120, 22) _
        .Items("Alpha", "Beta", "Gamma").Value(1).WritesTo "nativePick"
    app.Slider("volume").AtRect(160, 212, 130, 16).SliderRange(0, 100, 5) _
        .Value(35).WritesTo "volume"

    app.Label("lblInput").AtRect(24, 246, 200, 16).Text("Cell-backed input").Bold
    app.TextInput("username").At("C17").WritesTo "userName"
    app.Label("inputHint").AtRect(24, 290, 260, 16) _
        .Text("Type in the framed cell and press Enter.")

    app.Label("lblProgress").AtRect(24, 322, 200, 16).Text("Progress").Bold
    app.ProgressBar("meter").AtRect(24, 344, 240, 12).BindValue "volume"
    app.Label("meterLbl").AtRect(276, 340, 200, 18) _
        .BindText "volume", "Bound to the slider: {0}"
    app.Spinner("spin").AtRect 490, 336, 26, 26

    app.Card("inspector").AtRect(24, 376, 560, 110).Text("State inspector")
    app.Label("inspectorBody").AtRect(36, 404, 536, 74).BindText "inspector"

    app.SetState "notifications", False
    app.SetState "audit", False
    app.SetState "region", "North"
    app.SetState "nativePick", "Alpha"
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
    app.OnStateChanged "audit", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "region", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "nativePick", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "volume", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "userName", "WidgetGallery.RefreshInspector"
    app.OnStateChanged "lastAction", "WidgetGallery.RefreshInspector"
End Sub

Public Sub RefreshInspector()
    Dim app As ReDimUI
    Dim summary As String

    Set app = GalleryApp()
    summary = "notifications = " & CStr(app.State("notifications")) & _
        "   audit = " & CStr(app.State("audit")) & _
        "   region = " & CStr(app.State("region")) & _
        "   native = " & CStr(app.State("nativePick")) & vbLf & _
        "volume = " & CStr(app.State("volume")) & _
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
