Attribute VB_Name = "WidgetGallery"
' ReDim 0.20.0 (2026-09-22)
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
    Dim ui As ReDimUI
    Dim host As Worksheet

    Set host = ThisWorkbook.Worksheets(1)
    Set ui = ReDimUI.Mount(host, APP_ID)
    ui.ProtectSurface False
    ui.PrepareCanvas

    ui.Label("title").AtRect(24, 16, 300, 30).Text("Widget Gallery") _
        .FontSize(20).Bold

    ui.Label("lblButtons").AtRect(24, 60, 200, 16).Text("Buttons").Bold
    ui.Button("primary").AtRect(24, 80, 96, 30).Text("Primary").Primary _
        .OnClick "WidgetGallery.HandlePing"
    ui.Button("secondary").AtRect(128, 80, 96, 30).Text("Secondary") _
        .Secondary.OnClick "WidgetGallery.HandlePing"
    ui.Button("success").AtRect(232, 80, 96, 30).Text("Success").Success _
        .OnClick "WidgetGallery.HandleToastSuccess"
    ui.Button("danger").AtRect(336, 80, 96, 30).Text("Danger").Danger _
        .OnClick "WidgetGallery.HandleModal"
    ui.Button("busywork").AtRect(440, 80, 120, 30).Text("Run async work") _
        .OnClickAsync "WidgetGallery.SimulatedWork"
    ui.Label("oplog").AtRect(440, 114, 220, 16).BindText "oplog"

    ui.Label("lblValues").AtRect(24, 128, 200, 16).Text("Value controls").Bold
    ui.Toggle("notify").AtRect(24, 150, 44, 22).WritesTo "notifications"
    ui.Label("notifyLbl").AtRect(76, 152, 110, 18).Text("Notifications")
    ui.SelectBox("region").AtRect(196, 148, 130, 24) _
        .Items("North", "South", "East", "West").Value(1).WritesTo "region"
    ui.SelectBox("region").OnChange "WidgetGallery.HandleRegionChange"

    ui.Label("lblDrawn").AtRect(24, 188, 300, 16) _
        .Text("Drawn controls (fully themed)").Bold
    ui.TickBox("consent").AtRect(24, 212, 140, 18).Text("Log activity") _
        .WritesTo "drawnCheck"
    ui.RadioGroup("priority").AtRect(190, 204, 130, 60) _
        .Items("Low", "Medium", "High").Value(2).WritesTo "priority"
    ui.Stepper("volume").AtRect(350, 206, 120, 24).SliderRange(0, 100, 5) _
        .Value(35).WritesTo("volume").BindValue "volume"

    ui.Label("lblInput").AtRect(24, 280, 200, 16).Text("Cell-free fields").Bold
    ui.TextInput("username").AtRect(24, 304, 150, 22).WritesTo "userName"
    ui.ComboBox("fruit").AtRect(196, 304, 150, 22).WritesTo "fruit"
    ui.ComboBox("fruit").Items "Apple", "Apricot", "Avocado", "Banana", _
        "Blackberry", "Blueberry", "Cherry", "Coconut", "Cranberry", _
        "Date", "Fig", "Grape", "Grapefruit", "Guava", "Kiwi", "Lemon", _
        "Lime", "Mango", "Nectarine", "Orange"
    ui.TextInput("notes").AtRect(360, 288, 150, 44).MultiLine.WritesTo "notes"
    ui.Label("inputHint").AtRect(24, 336, 520, 16) _
        .Text("Click a field and type: the combo filters live; multi-line notes commit with Tab or Ctrl+Enter.")

    ui.Label("lblTransfer").AtRect(24, 356, 300, 16) _
        .Text("Transfer list (dual listbox)").Bold
    ui.Label("lblCheck").AtRect(420, 356, 150, 16) _
        .Text("Checkbox list").Bold
    ui.CheckList("options").AtRect 420, 378, 160, 125
    ui.CheckList("options").ItemsFrom( _
        Array("Alerts", "Auto-save", "Dark mode", "Sync")) _
        .CheckedFrom(Array("Auto-save")) _
        .WritesTo "options"
    ui.TransferList("crew").AtRect 24, 378, 380, 132
    ui.TransferList("crew").ItemsFrom( _
        Array("Ada", "Grace", "Edsger", "Alan", "Donald", "Katherine", _
            "Margaret", "John", "Dennis", "Ken", "Bjarne", "Linus", _
            "Guido", "Tim")) _
        .ChosenFrom(Array("Barbara")) _
        .Captions("Available", "On mission") _
        .WritesTo "crew"
    ui.TransferList("crew").OnChange "WidgetGallery.HandleCrewChange"

    ui.Label("lblImage").AtRect(420, 507, 150, 14).Text("Image").Bold
    ui.Image("logo").AtRect 420, 524, 160, 74
    ui.Image("logo").Source EnsureDemoImage(host)

    ui.Label("lblProgress").AtRect(24, 526, 200, 16) _
        .Text("Slider, stepper, and meter share one state key").Bold
    ui.SlideBar("volumeslide").AtRect(24, 548, 240, 18) _
        .SliderRange(0, 100, 5).Value(35).WritesTo("volume").BindValue "volume"
    ui.ProgressBar("meter").AtRect(24, 574, 240, 12).BindValue "volume"
    ui.Label("meterLbl").AtRect(276, 568, 220, 18) _
        .BindText "volume", "Slide, step, or watch: {0}"
    ui.Spinner("spin").AtRect 490, 564, 26, 26

    ui.Card("inspector").AtRect(24, 606, 560, 110).Text("State inspector")
    ui.Label("inspectorBody").AtRect(36, 634, 536, 74).BindText "inspector"

    ui.SetState "notifications", False
    ui.SetState "region", "North"
    ui.SetState "drawnCheck", False
    ui.SetState "priority", "Medium"
    ui.SetState "volume", 35
    ui.SetState "userName", vbNullString
    ui.SetState "fruit", vbNullString
    ui.SetState "notes", vbNullString
    ui.SetState "crew", "Barbara"
    ui.SetState "options", "Auto-save"
    ui.SetState "lastAction", "none yet"
    ui.SetState "oplog", "no run yet"
    RefreshInspector
    WireInspector ui
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub WireInspector(ByVal ui As ReDimUI)
    ui.OnStateChanged "notifications", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "region", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "drawnCheck", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "priority", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "volume", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "userName", "WidgetGallery.RefreshInspector"
    ui.OnStateChanged "lastAction", "WidgetGallery.RefreshInspector"
End Sub

Public Sub RefreshInspector()
    Dim ui As ReDimUI
    Dim inspectorText As String

    Set ui = GalleryApp()
    inspectorText = "notifications = " & CStr(ui.State("notifications")) & _
        "   region = " & CStr(ui.State("region")) & _
        "   check = " & CStr(ui.State("drawnCheck")) & vbLf & _
        "volume = " & CStr(ui.State("volume")) & _
        "   priority = " & CStr(ui.State("priority")) & _
        "   userName = " & CStr(ui.State("userName")) & vbLf & _
        "last action = " & CStr(ui.State("lastAction"))
    ui.SetState "inspector", inspectorText
End Sub

Public Sub HandleRegionChange()
    GalleryApp().SetState "lastAction", "region picked"
End Sub

Public Sub HandleCrewChange()
    GalleryApp().SetState "lastAction", "crew transferred"
End Sub

' A stand-in logo generated on the spot: overlapping shapes and text
' composed on a chart canvas and exported as a PNG. A solid-color image
' is indistinguishable from a plain filled shape, so the logo layers
' elements no shape fill could fake.
Private Function EnsureDemoImage(ByVal host As Worksheet) As String
    Dim chartHost As ChartObject
    Dim canvas As Chart
    Dim targetPath As String

    targetPath = ROneCOne.Path.Combine( _
        ROneCOne.Path.GetTempPath(), "rdm_gallery_logo.png")
    If ROneCOne.File.Exists(targetPath) Then ROneCOne.File.Delete targetPath
    Set chartHost = host.ChartObjects.Add(0, 0, 160, 74)
    Set canvas = chartHost.Chart
    canvas.ChartArea.Format.Fill.ForeColor.RGB = RGB(24, 56, 42)
    With canvas.Shapes.AddShape(msoShapeOval, 8, 10, 54, 54)
        .Fill.ForeColor.RGB = RGB(31, 111, 76)
        .Line.Visible = msoFalse
    End With
    With canvas.Shapes.AddShape(msoShapeOval, 34, 22, 34, 34)
        .Fill.ForeColor.RGB = RGB(240, 178, 54)
        .Line.Visible = msoFalse
    End With
    With canvas.Shapes.AddTextbox(msoTextOrientationHorizontal, 72, 22, 84, 30)
        .TextFrame2.TextRange.Text = "ReDim"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    canvas.Export targetPath, "PNG"
    chartHost.Delete
    EnsureDemoImage = targetPath
End Function

Public Sub HandlePing()
    Dim ui As ReDimUI

    Set ui = GalleryApp()
    ui.SetState "lastAction", ReDimUI.SenderId & " clicked"
    ui.Toast ReDimUI.SenderId & " clicked.", 1800
End Sub

Public Sub HandleToastSuccess()
    Dim ui As ReDimUI

    Set ui = GalleryApp()
    ui.SetState "lastAction", "success toast"
    ui.Toast "Everything saved cleanly.", 2500
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
    Dim ui As ReDimUI
    Dim ignored As Variant
    Dim startedAt As Double

    Set ui = GalleryApp()
    startedAt = Timer
    ui.SetState "lastAction", "async work running"
    ui.SetState "oplog", "op started " & Format$(startedAt, "0.0") & "s"
    ignored = ROneCOne.Task.Delay(1200).Await
    ui.SetState "lastAction", "async work finished"
    ui.SetState "oplog", "op ran " & Format$(Timer - startedAt, "0.00") & _
        "s (expected 1.2)"
End Sub
