Attribute VB_Name = "Navigator"
Option Explicit

' Navigator: sheets as forms. Three windows (Home, Settings, About), each a
' ReDim app on its own sheet, with Navigate, a back stack, OnShow lifecycle
' hooks, and persisted settings. Only the active window's sheet is visible;
' the others are very-hidden like closed forms.

Private gHomeShownCount As Long

Public Sub Auto_Open()
    BuildNavigator
    ReDimUI.Navigate "navhome"
End Sub

Public Sub BuildNavigator()
    BuildHome
    BuildSettings
    BuildAbout
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In ThisWorkbook.Worksheets
        If candidate.Name = sheetName Then
            Set EnsureSheet = candidate
            Exit Function
        End If
    Next candidate
    Set EnsureSheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    EnsureSheet.Name = sheetName
End Function

Private Sub BuildHome()
    Dim app As ReDimUI

    Set app = ReDimUI.Mount(EnsureSheet("NavHome"), "navhome")
    app.ProtectSurface False
    app.PrepareCanvas
    app.AsWindow.WindowTitle "Home"
    app.OnShow "Navigator.HandleHomeShown"
    app.NavBar

    app.Label("title").AtRect(24, 48, 320, 30).Text("Navigator").FontSize(20).Bold
    app.Label("subtitle").Below("title", 2).Sized(420, 18) _
        .Text("Sheets as forms: tabs navigate, back walks the stack.")
    app.Label("visits").Below("subtitle", 16).Sized(300, 18) _
        .BindText "homeShown", "Home shown {0} times this session."
    app.SetState "homeShown", 0
    app.Render
    app.ProtectSurface
End Sub

Private Sub BuildSettings()
    Dim app As ReDimUI

    Set app = ReDimUI.Mount(EnsureSheet("NavSettings"), "navsettings")
    app.ProtectSurface False
    app.PrepareCanvas
    app.AsWindow.WindowTitle "Settings"
    app.Persist True
    app.NavBar

    app.Label("title").AtRect(24, 48, 320, 30).Text("Settings").FontSize(20).Bold
    app.Toggle("alerts").Below("title", 16).Sized(44, 22).WritesTo "alertsOn"
    app.Label("alertslbl").RightOf("alerts", 10).Sized(200, 18) _
        .Text("Alert notifications")
    app.Toggle("autosave").Below("alerts", 14).Sized(44, 22).WritesTo "autoSave"
    app.Label("autosavelbl").RightOf("autosave", 10).Sized(200, 18) _
        .Text("Autosave results")
    app.Label("hint").Below("autosave", 20).Sized(360, 18) _
        .Text("These choices persist across closing the workbook.")
    app.Button("back").Below("hint", 14).Sized(110, 30).Text("< Back") _
        .Secondary.OnClick "Navigator.HandleBack"
    app.SetStateDefault "alertsOn", True
    app.SetStateDefault "autoSave", False
    app.Toggle("alerts").Checked CBool(app.State("alertsOn"))
    app.Toggle("autosave").Checked CBool(app.State("autoSave"))
    app.Render
    app.ProtectSurface
End Sub

Private Sub BuildAbout()
    Dim app As ReDimUI

    Set app = ReDimUI.Mount(EnsureSheet("NavAbout"), "navabout")
    app.ProtectSurface False
    app.PrepareCanvas
    app.AsWindow.WindowTitle "About"
    app.NavBar

    app.Label("title").AtRect(24, 48, 320, 30).Text("About").FontSize(20).Bold
    app.Card("card").Below("title", 14).Sized(380, 110) _
        .Text("ReDim " & ReDimUI.Version & vbLf & vbLf & _
            "A stateful UI framework for Excel worksheets, built on ROneCOne." & _
            vbLf & "github.com/WilliamSmithEdward/ReDim")
    app.Button("back").Below("card", 14).Sized(110, 30).Text("< Back") _
        .Secondary.OnClick "Navigator.HandleBack"
    app.Render
    app.ProtectSurface
End Sub

Public Sub HandleHomeShown()
    Dim app As ReDimUI

    gHomeShownCount = gHomeShownCount + 1
    Set app = ReDimUI.App("navhome")
    app.SetState "homeShown", gHomeShownCount
End Sub

Public Sub HandleBack()
    If Not ReDimUI.NavigateBack() Then ReDimUI.Navigate "navhome"
End Sub
