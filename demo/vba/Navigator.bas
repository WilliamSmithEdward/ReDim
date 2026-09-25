Attribute VB_Name = "Navigator"
' ReDim 1.0.2 (2026-09-23)
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
    Dim ui As ReDimUI

    Set ui = ReDimUI.Mount(EnsureSheet("NavHome"), "navhome")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.AsWindow.WindowTitle "Home"
    ui.OnShow "Navigator.HandleHomeShown"
    ui.NavBar

    ui.Label("title").AtRect(24, 48, 320, 30).Text("Navigator").FontSize(20).Bold
    ui.Label("subtitle").Below("title", 2).Sized(420, 18) _
        .Text("Sheets as forms: tabs navigate, back walks the stack.")
    ui.Label("visits").Below("subtitle", 16).Sized(300, 18) _
        .BindText "homeShown", "Home shown {0} times this session."
    ui.SetState "homeShown", 0
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub BuildSettings()
    Dim ui As ReDimUI

    Set ui = ReDimUI.Mount(EnsureSheet("NavSettings"), "navsettings")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.AsWindow.WindowTitle "Settings"
    ui.NavBar

    ui.Label("title").AtRect(24, 48, 320, 30).Text("Settings").FontSize(20).Bold
    ui.Toggle("alerts").Below("title", 16).Sized(44, 22).Text("Alert notifications") _
        .WritesTo "alertsOn"
    ui.Toggle("autosave").Below("alerts", 14).Sized(44, 22).Text("Autosave results") _
        .WritesTo "autoSave"
    ui.Label("hint").Below("autosave", 20).Sized(360, 18) _
        .Text("These choices live in app state for this session.")
    ui.Button("back").Below("hint", 14).Sized(110, 30).Text("< Back") _
        .Secondary.OnClick "Navigator.HandleBack"
    ' Each switch follows the key it writes, so the defaults show on it.
    ui.SetStateDefault "alertsOn", True
    ui.SetStateDefault "autoSave", False
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub BuildAbout()
    Dim ui As ReDimUI

    Set ui = ReDimUI.Mount(EnsureSheet("NavAbout"), "navabout")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.AsWindow.WindowTitle "About"
    ui.NavBar

    ui.Label("title").AtRect(24, 48, 320, 30).Text("About").FontSize(20).Bold
    ui.Card("card").Below("title", 14).Sized(380, 110) _
        .Text("ReDim " & ReDimUI.Version & vbLf & vbLf & _
            "A stateful UI framework for Excel worksheets, built on ROneCOne." & _
            vbLf & "github.com/WilliamSmithEdward/ReDim")
    ui.Button("back").Below("card", 14).Sized(110, 30).Text("< Back") _
        .Secondary.OnClick "Navigator.HandleBack"
    ui.Render
    ui.ProtectSurface
End Sub

Public Sub HandleHomeShown()
    Dim ui As ReDimUI

    gHomeShownCount = gHomeShownCount + 1
    Set ui = ReDimUI.App("navhome")
    ui.SetState "homeShown", gHomeShownCount
End Sub

Public Sub HandleBack()
    If Not ReDimUI.NavigateBack() Then ReDimUI.Navigate "navhome"
End Sub
