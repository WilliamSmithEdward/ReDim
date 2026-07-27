Attribute VB_Name = "PokeDex"
Option Explicit

' ReDex: a living Pokedex in Excel, and the full-framework showcase.
' Live PokeAPI data rides ROneCOne HTTP tasks through the pump, sprites
' download into Image controls, stat bars tween through a paced job,
' three windows navigate under a NavBar, and every drawn control in the
' family appears somewhere it genuinely belongs. Entry: BuildPokeDex.

Private Declare PtrSafe Function URLDownloadToFileA Lib "urlmon" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As LongPtr _
) As Long

Private Const API_BASE As String = "https://pokeapi.co/api/v2/pokemon/"
Private Const DEX_MAX As Long = 151

' Task references live here so done handlers can read Result; the op
' itself keeps no public handle on its task.
Private gNamesTask As ROneCOne
Private gDetailTask As ROneCOne
' The stat tween model: bars ease from now toward target every frame.
Private gStatNow(1 To 6) As Double
Private gStatTarget(1 To 6) As Double
Private gHuntStep As Long

Public Sub BuildPokeDex()
    BuildBrowse
    BuildTeam
    BuildTrainer
    ReDimUI.Navigate "dexbrowse"
End Sub

' =====================================================================
' Windows
' =====================================================================

Private Sub BuildBrowse()
    Dim app As ReDimUI
    Dim host As Worksheet
    Dim rowIndex As Long

    Set host = EnsureSheet("DexBrowse")
    Set app = ReDimUI.Mount(host, "dexbrowse")
    app.ProtectSurface False
    app.PrepareCanvas
    app.SetTheme PokeTheme(False)
    app.AsWindow.WindowTitle "Pokedex"
    app.NavBar
    app.OnShow "PokeDex.HandleBrowseShown"

    app.Label("title").AtRect(24, 46, 220, 34).Text("ReDex").FontSize(24).Bold
    app.Label("tag").AtRect(24, 82, 400, 16) _
        .Text("A living Pokedex, drawn from shapes, fed by PokeAPI.")

    app.ComboBox("species").AtRect(24, 108, 200, 24).WritesTo "speciesPick"
    app.ComboBox("species").OnChange "PokeDex.HandleSpeciesPick"
    app.Button("prev").AtRect(240, 108, 70, 24).Text("< F2").Secondary _
        .OnClick "PokeDex.GoPrev"
    app.Button("nextb").AtRect(318, 108, 70, 24).Text("F3 >").Secondary _
        .OnClick "PokeDex.GoNext"
    app.Spinner("spn").AtRect 400, 108, 22, 22
    app.Label("status").AtRect(432, 112, 170, 16).BindText "fetchStatus"

    app.Image("sprite").AtRect(24, 148, 150, 150).BindSource "spriteFile"
    app.Card("infocard").AtRect(190, 148, 412, 150).Text("")
    app.Label("pokename").AtRect(206, 160, 280, 24).FontSize(15).Bold _
        .BindText "pokeName"
    app.Label("typea").AtRect(206, 192, 88, 20).BindText("typeA") _
        .BindVisible "typeAOn"
    app.Label("typeb").AtRect(302, 192, 88, 20).BindText("typeB") _
        .BindVisible "typeBOn"
    app.Label("sizes").AtRect(206, 222, 380, 16).BindText "sizeLine"
    app.Label("dexline").AtRect(206, 244, 380, 16).BindText "dexLine"

    app.Label("statshdr").AtRect(24, 314, 200, 16).Text("Base stats").Bold
    For rowIndex = 1 To 6
        app.Label("statn" & rowIndex) _
            .AtRect(24, 320 + rowIndex * 22, 64, 16).Text(StatCaption(rowIndex))
        app.ProgressBar("statb" & rowIndex) _
            .AtRect(94, 323 + rowIndex * 22, 220, 10).BindValue StatPctKey(rowIndex)
        app.Label("statv" & rowIndex) _
            .AtRect(322, 320 + rowIndex * 22, 44, 16).BindText StatKey(rowIndex)
    Next rowIndex

    app.Button("catch").AtRect(400, 342, 202, 34).Text("Catch!").Primary _
        .OnClick "PokeDex.CatchCurrent"
    app.Button("hunt").AtRect(400, 386, 130, 26).Text("Shiny hunt") _
        .Secondary.OnClick "PokeDex.StartShinyHunt"
    app.Button("stophunt").AtRect(538, 386, 64, 26).Text("Stop").Danger _
        .OnClick "PokeDex.StopShinyHunt"
    app.Button("hunt").BindEnabled "hunting", True
    app.Button("stophunt").BindEnabled "hunting"
    app.ProgressBar("huntbar").AtRect(400, 420, 202, 8).BindValue "huntPct"

    app.SetStateDefault "fetchStatus", "idle"
    app.SetStateDefault "pokeName", "Loading the Kanto dex..."
    app.SetStateDefault "typeAOn", False
    app.SetStateDefault "typeBOn", False
    app.SetStateDefault "sizeLine", ""
    app.SetStateDefault "dexLine", ""
    app.SetStateDefault "spriteFile", ""
    app.SetStateDefault "hunting", False
    app.SetStateDefault "huntPct", 0
    app.SetStateDefault "tweening", False
    app.SetStateDefault "namesLoaded", False
    app.SetStateDefault "dexId", 0
    app.HotKey "{F2}", "PokeDex.GoPrev"
    app.HotKey "{F3}", "PokeDex.GoNext"
    app.Render
    app.ProtectSurface
End Sub

Private Sub BuildTeam()
    Dim app As ReDimUI
    Dim host As Worksheet

    Set host = EnsureSheet("DexTeam")
    Set app = ReDimUI.Mount(host, "dexteam")
    app.ProtectSurface False
    app.PrepareCanvas
    app.SetTheme PokeTheme(False)
    app.AsWindow.WindowTitle "Team"
    app.NavBar
    app.OnShow "PokeDex.HandleTeamShown"

    app.Label("title").AtRect(24, 46, 300, 26).Text("Team builder") _
        .FontSize(18).Bold
    app.Label("hint").AtRect(24, 76, 420, 16) _
        .Text("Catch species in the Pokedex, then build a party of six.")
    app.TransferList("team").AtRect 24, 100, 420, 152
    app.TransferList("team").Captions("Caught", "Party") _
        .WritesTo("party").OnChange "PokeDex.HandlePartyChange"

    app.Label("preflbl").AtRect(470, 76, 150, 16).Text("House rules").Bold
    app.CheckList("prefs").AtRect 470, 100, 160, 152
    app.CheckList("prefs").ItemsFrom( _
        Array("Nicknames", "Auto-heal", "Hard mode", "Shiny only")) _
        .CheckedFrom(Array("Auto-heal")) _
        .WritesTo "houseRules"

    app.Label("noteslbl").AtRect(24, 268, 200, 16).Text("Strategy notes").Bold
    app.TextInput("strategy").AtRect(24, 288, 300, 58).MultiLine _
        .WritesTo "strategy"
    app.Label("partylbl").AtRect(340, 288, 290, 58) _
        .BindText "party", "Party: {0}"
    app.Label("visits").AtRect(24, 360, 300, 16).BindText _
        "teamVisits", "Window shown {0} times this session."

    app.SetStateDefault "party", ""
    app.SetStateDefault "strategy", ""
    app.SetStateDefault "teamVisits", 0
    app.Render
    app.ProtectSurface
End Sub

Private Sub BuildTrainer()
    Dim app As ReDimUI
    Dim host As Worksheet

    Set host = EnsureSheet("DexTrainer")
    Set app = ReDimUI.Mount(host, "dextrainer")
    app.ProtectSurface False
    app.PrepareCanvas
    app.SetTheme PokeTheme(False)
    app.AsWindow.WindowTitle "Trainer"
    app.NavBar
    app.OnShow "PokeDex.RefreshTrainerCard"

    app.Label("title").AtRect(24, 46, 300, 26).Text("Trainer card") _
        .FontSize(18).Bold
    app.Label("namelbl").AtRect(24, 84, 120, 18).Text("Name")
    app.TextInput("trainer").AtRect(120, 82, 170, 22).WritesTo "trainerName"
    app.Label("starterlbl").AtRect(24, 118, 120, 18).Text("Starter").Bold
    app.RadioGroup("starter").AtRect(24, 140, 160, 62) _
        .Items("Bulbasaur", "Charmander", "Squirtle").Value(1) _
        .WritesTo("starter").OnChange "PokeDex.HandleStarterPick"
    app.Label("favlbl").AtRect(220, 118, 120, 18).Text("Favorite type").Bold
    app.SelectBox("favtype").AtRect(220, 140, 140, 24) _
        .Items("Fire", "Water", "Grass", "Electric", "Dragon") _
        .Value(2).WritesTo "favType"
    app.Label("lvllbl").AtRect(220, 178, 120, 18).Text("Ambition (level)")
    app.Stepper("level").AtRect(220, 198, 140, 24).SliderRange(1, 100, 5) _
        .Value(5).WritesTo "ambition"

    app.Toggle("darkmode").AtRect(24, 226, 44, 22).WritesTo("darkMode") _
        .OnChange "PokeDex.ApplyThemeChoice"
    app.Label("darklbl").AtRect(76, 228, 160, 18).Text("Night mode")

    app.Card("card").AtRect(24, 262, 420, 96).Text("Trainer summary")
    app.Label("cardbody").AtRect(36, 290, 396, 60).BindText "cardText"
    app.Button("reset").AtRect(470, 262, 150, 30).Text("Reset journey") _
        .Danger.OnClick "PokeDex.ConfirmReset"

    app.SetStateDefault "trainerName", "Red"
    app.SetStateDefault "starter", "Bulbasaur"
    app.SetStateDefault "favType", "Water"
    app.SetStateDefault "ambition", 5
    app.SetStateDefault "darkMode", False
    app.SetStateDefault "cardText", ""
    app.OnStateChanged "trainerName", "PokeDex.RefreshTrainerCard"
    app.OnStateChanged "starter", "PokeDex.RefreshTrainerCard"
    app.OnStateChanged "favType", "PokeDex.RefreshTrainerCard"
    app.OnStateChanged "ambition", "PokeDex.RefreshTrainerCard"
    app.Render
    app.ProtectSurface
End Sub

' =====================================================================
' Fetching: names list once, then details per selection
' =====================================================================

Public Sub HandleBrowseShown()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    If CBool(app.StateOrDefault("namesLoaded", False)) Then Exit Sub
    If app.IsOpRunning("names") Then Exit Sub
    Set gNamesTask = ROneCOne.HttpClient.GetStringAsync( _
        API_BASE & "?limit=" & DEX_MAX)
    With app.Async("names")
        .RunsTask gNamesTask
        .ShowsSpinner "spn"
        .TracksState "fetchStatus"
        .OnDone "PokeDex.ApplyNamesList"
        .OnFail "PokeDex.NamesFailed"
    End With
    app.Async("names").Start
End Sub

Public Sub ApplyNamesList()
    Dim app As ReDimUI
    Dim doc As ROneCOne
    Dim resultsList As ROneCOne
    Dim entry As ROneCOne
    Dim names As Collection
    Dim position As Long

    Set app = ReDimUI.App("dexbrowse")
    Set doc = ROneCOne.Json.Deserialize(CStr(gNamesTask.Result))
    Set resultsList = doc.Item("results")
    Set names = New Collection
    For position = 0 To resultsList.Count - 1
        Set entry = resultsList.Item(position)
        names.Add StrConv(CStr(entry.Item("name")), vbProperCase)
    Next position
    app.ComboBox("species").ItemsFrom names
    app.SetState "namesLoaded", True
    app.Toast names.Count & " Kanto species loaded. Pick one!", 3500
    FetchByPath "1"
End Sub

Public Sub NamesFailed()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    app.SetState "pokeName", "PokeAPI unreachable."
    app.Toast "Could not reach PokeAPI: " & app.AsyncError("names"), 6000
End Sub

Public Sub HandleSpeciesPick()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    FetchByPath LCase$(CStr(app.StateOrDefault("speciesPick", "")))
End Sub

Public Sub GoPrev()
    StepDex -1
End Sub

Public Sub GoNext()
    StepDex 1
End Sub

Private Sub StepDex(ByVal delta As Long)
    Dim app As ReDimUI
    Dim dexId As Long

    Set app = ReDimUI.App("dexbrowse")
    dexId = CLng(app.StateOrDefault("dexId", 0)) + delta
    If dexId < 1 Then dexId = DEX_MAX
    If dexId > DEX_MAX Then dexId = 1
    FetchByPath CStr(dexId)
End Sub

Private Sub FetchByPath(ByVal pathPart As String)
    Dim app As ReDimUI

    If LenB(pathPart) = 0 Then Exit Sub
    Set app = ReDimUI.App("dexbrowse")
    If app.IsOpRunning("detail") Then Exit Sub
    Set gDetailTask = ROneCOne.HttpClient.GetStringAsync(API_BASE & pathPart)
    With app.Async("detail")
        .RunsTask gDetailTask
        .ShowsSpinner "spn"
        .TracksState "fetchStatus"
        .OnDone "PokeDex.ApplyDetails"
        .OnFail "PokeDex.DetailFailed"
    End With
    app.Async("detail").Start
End Sub

Public Sub ApplyDetails()
    Dim app As ReDimUI
    Dim doc As ROneCOne
    Dim typesList As ROneCOne
    Dim statsList As ROneCOne
    Dim entry As ROneCOne
    Dim inner As ROneCOne
    Dim displayName As String
    Dim dexId As Long
    Dim position As Long
    Dim statIndex As Long

    Set app = ReDimUI.App("dexbrowse")
    Set doc = ROneCOne.Json.Deserialize(CStr(gDetailTask.Result))
    dexId = CLng(doc.Item("id"))
    displayName = StrConv(CStr(doc.Item("name")), vbProperCase)

    app.BeginUpdate
    app.SetState "dexId", dexId
    app.SetState "pokeName", "#" & Format$(dexId, "000") & "  " & displayName
    app.SetState "sizeLine", _
        Format$(CDbl(doc.Item("height")) / 10, "0.0") & " m  /  " & _
        Format$(CDbl(doc.Item("weight")) / 10, "0.0") & " kg"
    app.SetState "dexLine", "Kanto dex " & dexId & " of " & DEX_MAX

    Set typesList = doc.Item("types")
    ApplyTypeBadge app, "typea", "typeA", typesList, 0
    ApplyTypeBadge app, "typeb", "typeB", typesList, 1

    Set statsList = doc.Item("stats")
    For position = 0 To statsList.Count - 1
        Set entry = statsList.Item(position)
        Set inner = entry.Item("stat")
        statIndex = StatIndexOf(CStr(inner.Item("name")))
        If statIndex > 0 Then
            gStatTarget(statIndex) = CDbl(entry.Item("base_stat"))
        End If
    Next position
    app.EndUpdate
    StartStatTween

    app.ComboBox("species").InputValue = displayName
    ApplySprite app, doc, dexId
End Sub

Public Sub DetailFailed()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    app.Toast "Fetch failed: " & app.AsyncError("detail"), 5000
End Sub

Private Sub ApplyTypeBadge( _
    ByVal app As ReDimUI, _
    ByVal labelId As String, _
    ByVal stateKey As String, _
    ByVal typesList As ROneCOne, _
    ByVal slotIndex As Long _
)
    Dim entry As ROneCOne
    Dim inner As ROneCOne
    Dim typeName As String

    If slotIndex < typesList.Count Then
        Set entry = typesList.Item(slotIndex)
        Set inner = entry.Item("type")
        typeName = CStr(inner.Item("name"))
        app.Label(labelId).Fill(TypeColor(typeName)) _
            .TextColor RGB(255, 255, 255)
        app.SetState stateKey, "  " & StrConv(typeName, vbProperCase)
        app.SetState stateKey & "On", True
    Else
        app.SetState stateKey & "On", False
    End If
End Sub

Private Sub ApplySprite( _
    ByVal app As ReDimUI, _
    ByVal doc As ROneCOne, _
    ByVal dexId As Long _
)
    Dim sprites As ROneCOne
    Dim spriteUrl As Variant
    Dim localPath As String

    Set sprites = doc.Item("sprites")
    spriteUrl = sprites.Item("front_default")
    If IsNull(spriteUrl) Or LenB(CStr(spriteUrl)) = 0 Then Exit Sub
    localPath = Environ$("TEMP") & "\redex_" & dexId & ".png"
    If LenB(Dir$(localPath)) = 0 Then
        If URLDownloadToFileA(0, CStr(spriteUrl), localPath, 0, 0) <> 0 Then
            app.Toast "Sprite download failed.", 4000
            Exit Sub
        End If
    End If
    app.SetState "spriteFile", localPath
End Sub

' =====================================================================
' Stat tween: bars ease toward their targets on a 16 ms paced job
' =====================================================================

Private Sub StartStatTween()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    If CBool(app.StateOrDefault("tweening", False)) Then Exit Sub
    app.SetState "tweening", True
    app.Job("tween").Steps("PokeDex.TweenStats").PacedMs 16
    app.Job("tween").StartJob
End Sub

Public Function TweenStats() As Boolean
    Dim app As ReDimUI
    Dim statIndex As Long
    Dim allDone As Boolean

    Set app = ReDimUI.App("dexbrowse")
    allDone = True
    app.BeginUpdate
    For statIndex = 1 To 6
        gStatNow(statIndex) = gStatNow(statIndex) + _
            (gStatTarget(statIndex) - gStatNow(statIndex)) * 0.22
        If Abs(gStatTarget(statIndex) - gStatNow(statIndex)) > 0.8 Then
            allDone = False
        Else
            gStatNow(statIndex) = gStatTarget(statIndex)
        End If
        app.SetState StatKey(statIndex), CLng(gStatNow(statIndex))
        app.SetState StatPctKey(statIndex), _
            CLng(gStatNow(statIndex) / 1.6)
    Next statIndex
    app.EndUpdate
    If allDone Then app.SetState "tweening", False
    TweenStats = allDone
End Function

' =====================================================================
' Catching, shiny hunting, party rules
' =====================================================================

Public Sub CatchCurrent()
    Dim browseApp As ReDimUI
    Dim teamApp As ReDimUI
    Dim caughtName As String
    Dim position As Long

    Set browseApp = ReDimUI.App("dexbrowse")
    caughtName = CStr(browseApp.StateOrDefault("pokeName", ""))
    If InStr(caughtName, "  ") = 0 Then Exit Sub
    caughtName = Mid$(caughtName, InStr(caughtName, "  ") + 2)
    Set teamApp = ReDimUI.App("dexteam")
    For position = 1 To teamApp.TransferList("team").ItemCount
        If teamApp.TransferList("team").ItemTextAt(position) = caughtName Then
            browseApp.Toast caughtName & " is already in your box.", 3000
            Exit Sub
        End If
    Next position
    teamApp.TransferList("team").AddItem caughtName
    browseApp.Toast "Gotcha! " & caughtName & " was caught!", 3500
End Sub

Public Sub StartShinyHunt()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    If CBool(app.StateOrDefault("hunting", False)) Then Exit Sub
    gHuntStep = 0
    app.SetState "hunting", True
    app.SetState "huntPct", 0
    app.Job("hunt").Steps("PokeDex.HuntStep").PacedMs 60
    app.Job("hunt").JobOnDone "PokeDex.HuntDone"
    app.Job("hunt").StartJob
End Sub

Public Function HuntStep() As Boolean
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    gHuntStep = gHuntStep + 1
    app.SetState "huntPct", gHuntStep
    HuntStep = (gHuntStep >= 100)
End Function

Public Sub HuntDone()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    app.SetState "hunting", False
    app.SetState "huntPct", 0
    If CLng(app.StateOrDefault("dexId", 0)) Mod 8 = 1 Then
        app.Toast "It sparkles... a SHINY appeared!", 6000
    Else
        app.Toast "No shiny this time. The hunt continues.", 4000
    End If
End Sub

Public Sub StopShinyHunt()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexbrowse")
    If Not CBool(app.StateOrDefault("hunting", False)) Then Exit Sub
    app.CancelJob "hunt"
    app.SetState "hunting", False
    app.SetState "huntPct", 0
    app.Toast "Hunt called off.", 3000
End Sub

Public Sub HandlePartyChange()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexteam")
    If app.TransferList("team").ChosenCount > 6 Then
        app.Toast "A party carries six! The rest ride in the box.", 4000
    End If
End Sub

Public Sub HandleTeamShown()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dexteam")
    app.SetState "teamVisits", CLng(app.StateOrDefault("teamVisits", 0)) + 1
End Sub

Public Sub HandleStarterPick()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dextrainer")
    app.Toast CStr(app.State("starter")) & ", I choose you!", 3000
    FetchByPath LCase$(CStr(app.State("starter")))
    ReDimUI.Navigate "dexbrowse"
End Sub

' =====================================================================
' Trainer card, theming, reset
' =====================================================================

Public Sub RefreshTrainerCard()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dextrainer")
    app.SetState "cardText", _
        "Trainer " & CStr(app.StateOrDefault("trainerName", "Red")) & _
        "   Starter: " & CStr(app.StateOrDefault("starter", "-")) & vbLf & _
        "Favorite type: " & CStr(app.StateOrDefault("favType", "-")) & _
        "   Aiming for level " & CStr(app.StateOrDefault("ambition", 5))
End Sub

Public Sub ApplyThemeChoice()
    Dim trainerApp As ReDimUI
    Dim darkMode As Boolean

    Set trainerApp = ReDimUI.App("dextrainer")
    darkMode = CBool(trainerApp.StateOrDefault("darkMode", False))
    ReDimUI.App("dexbrowse").SetTheme PokeTheme(darkMode)
    ReDimUI.App("dexteam").SetTheme PokeTheme(darkMode)
    trainerApp.SetTheme PokeTheme(darkMode)
    trainerApp.Toast IIf(darkMode, "Lights out.", "Rise and shine."), 2500
End Sub

Public Sub ConfirmReset()
    Dim app As ReDimUI

    Set app = ReDimUI.App("dextrainer")
    app.Confirm "Reset journey?", _
        "Your box, party, and notes go back to square one.", _
        "PokeDex.DoReset"
End Sub

Public Sub DoReset()
    Dim teamApp As ReDimUI

    Set teamApp = ReDimUI.App("dexteam")
    teamApp.TransferList("team").ClearItems
    teamApp.TransferList("team").ChosenFrom Array()
    teamApp.SetState "party", ""
    teamApp.SetState "strategy", ""
    teamApp.TextInput("strategy").InputValue = ""
    ReDimUI.App("dextrainer").Toast "A fresh journey begins.", 3500
End Sub

' =====================================================================
' Helpers
' =====================================================================

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In ActiveWorkbook.Worksheets
        If candidate.Name = sheetName Then
            Set EnsureSheet = candidate
            Exit Function
        End If
    Next candidate
    Set EnsureSheet = ActiveWorkbook.Worksheets.Add
    EnsureSheet.Name = sheetName
End Function

' A Pokedex-red theme in light and night variants, built on the same
' ConfigureTheme surface the stock themes use.
Private Function PokeTheme(ByVal darkMode As Boolean) As ReDimUI
    Dim themeValue As ReDimUI

    Set themeValue = New ReDimUI
    If darkMode Then
        themeValue.ConfigureTheme _
            RGB(232, 84, 74), RGB(20, 20, 22), _
            RGB(45, 44, 48), RGB(240, 238, 235), _
            RGB(66, 64, 70), RGB(176, 172, 168), _
            RGB(96, 200, 140), RGB(240, 110, 110), _
            RGB(92, 88, 94), RGB(30, 29, 33), _
            "Segoe UI", 11
    Else
        themeValue.ConfigureTheme _
            RGB(214, 55, 46), RGB(255, 255, 255), _
            RGB(255, 255, 255), RGB(40, 40, 42), _
            RGB(238, 233, 229), RGB(122, 116, 112), _
            RGB(46, 140, 90), RGB(178, 34, 52), _
            RGB(220, 212, 206), RGB(250, 247, 244), _
            "Segoe UI", 11
    End If
    Set PokeTheme = themeValue
End Function

Private Function StatCaption(ByVal statIndex As Long) As String
    Select Case statIndex
        Case 1: StatCaption = "HP"
        Case 2: StatCaption = "Attack"
        Case 3: StatCaption = "Defense"
        Case 4: StatCaption = "Sp. Atk"
        Case 5: StatCaption = "Sp. Def"
        Case 6: StatCaption = "Speed"
    End Select
End Function

Private Function StatKey(ByVal statIndex As Long) As String
    Select Case statIndex
        Case 1: StatKey = "statHp"
        Case 2: StatKey = "statAtk"
        Case 3: StatKey = "statDef"
        Case 4: StatKey = "statSpa"
        Case 5: StatKey = "statSpd"
        Case 6: StatKey = "statSpe"
    End Select
End Function

Private Function StatPctKey(ByVal statIndex As Long) As String
    StatPctKey = StatKey(statIndex) & "Pct"
End Function

Private Function StatIndexOf(ByVal apiName As String) As Long
    Select Case apiName
        Case "hp": StatIndexOf = 1
        Case "attack": StatIndexOf = 2
        Case "defense": StatIndexOf = 3
        Case "special-attack": StatIndexOf = 4
        Case "special-defense": StatIndexOf = 5
        Case "speed": StatIndexOf = 6
    End Select
End Function

' Canonical franchise type colors.
Private Function TypeColor(ByVal typeName As String) As Long
    Select Case typeName
        Case "normal": TypeColor = RGB(168, 168, 120)
        Case "fire": TypeColor = RGB(240, 128, 48)
        Case "water": TypeColor = RGB(104, 144, 240)
        Case "grass": TypeColor = RGB(120, 200, 80)
        Case "electric": TypeColor = RGB(216, 172, 26)
        Case "ice": TypeColor = RGB(122, 199, 193)
        Case "fighting": TypeColor = RGB(192, 48, 40)
        Case "poison": TypeColor = RGB(160, 64, 160)
        Case "ground": TypeColor = RGB(202, 166, 90)
        Case "flying": TypeColor = RGB(150, 138, 224)
        Case "psychic": TypeColor = RGB(248, 88, 136)
        Case "bug": TypeColor = RGB(150, 168, 34)
        Case "rock": TypeColor = RGB(184, 160, 56)
        Case "ghost": TypeColor = RGB(112, 88, 152)
        Case "dragon": TypeColor = RGB(112, 56, 248)
        Case "dark": TypeColor = RGB(112, 88, 72)
        Case "steel": TypeColor = RGB(148, 152, 168)
        Case "fairy": TypeColor = RGB(220, 130, 160)
        Case Else: TypeColor = RGB(120, 120, 120)
    End Select
End Function
