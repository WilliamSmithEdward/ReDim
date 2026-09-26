Attribute VB_Name = "PokeDex"
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

' ReDex: a living Pokedex in Excel, and the full-framework showcase.
' Live PokeAPI data rides ROneCOne HTTP tasks through the pump, sprites
' download into Image controls, stat bars tween through a paced job,
' three windows navigate under a NavBar, and every drawn control in the
' family appears somewhere it genuinely belongs. Entry: BuildPokeDex.

Private Const API_BASE As String = "https://pokeapi.co/api/v2/pokemon/"
Private Const DEX_MAX As Long = 151

' Task references live here so done handlers can read Result; the op
' itself keeps no public handle on its task.
Private gNamesTask As ROneCOne
Private gDetailTask As ROneCOne
' Latest-wins switching. gWantPath is the newest selection, gFlightPath
' the one the in-flight op is fetching, gShownPath the one on screen.
' Parsed documents cache by request key plus canonical id and name, so
' revisits apply with no HTTP and no parse at all, and the nav cursor
' advances on the keypress itself so held keys step through the dex.
Private gDetailCache As ROneCOne
Private gWantPath As String
Private gFlightPath As String
Private gShownPath As String
Private gNavId As Long
' Sprites ride the pump like every other fetch. The image is kept in
' temp and reused, so a species seen once paints from disk with no
' request at all, and a fast switch never blocks on a download.
Private gSpriteTask As ROneCOne
Private gSpriteWantId As Long
Private gSpriteWantUrl As String
Private gSpriteFlightId As Long
' The stat tween model: bars ease from now toward target every frame.
Private gStatNow(1 To 6) As Double
Private gStatTarget(1 To 6) As Double
Private gHuntStep As Long

' The workbook builds itself when a user opens it with macros enabled.
' Automation opens (the test harness) skip Auto_Open, so scenarios stay
' in control of when the build runs.
Public Sub Auto_Open()
    BuildPokeDex
End Sub

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
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim statRow As Long

    Set host = EnsureSheet("DexBrowse")
    Set ui = ReDimUI.Mount(host, "dexbrowse")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.SetTheme PokeTheme(False)
    ui.AsWindow.WindowTitle "Pokedex"
    ui.NavBar
    ui.OnShow "PokeDex.HandleBrowseShown"

    ui.Label("title").AtRect(24, 46, 220, 34).Text("ReDex").FontSize(24).Bold
    ui.Label("tag").AtRect(24, 82, 400, 16) _
        .Text("A living Pokedex, drawn from shapes, fed by PokeAPI.")

    ui.ComboBox("species").AtRect(24, 108, 200, 24).WritesTo "speciesPick"
    ui.ComboBox("species").OnChange "PokeDex.HandleSpeciesPick"
    ui.Button("prev").AtRect(240, 108, 70, 24).Text("< F2").Secondary _
        .OnClick "PokeDex.GoPrev"
    ui.Button("nextb").AtRect(318, 108, 70, 24).Text("F3 >").Secondary _
        .OnClick "PokeDex.GoNext"
    ui.Spinner("spn").AtRect 400, 108, 22, 22
    ui.Label("status").AtRect(432, 112, 170, 16).BindText "fetchStatus"

    ui.Image("sprite").AtRect(24, 148, 150, 150).BindSource "spriteFile"
    ui.Card("infocard").AtRect(190, 148, 412, 150).Text("")
    ui.Label("pokename").AtRect(206, 160, 280, 24).FontSize(15).Bold _
        .BindText "pokeName"
    ui.Label("typea").AtRect(206, 192, 88, 20).BindText("typeA") _
        .BindVisible "typeAOn"
    ui.Label("typeb").AtRect(302, 192, 88, 20).BindText("typeB") _
        .BindVisible "typeBOn"
    ui.Label("sizes").AtRect(206, 222, 380, 16).BindText "sizeLine"
    ui.Label("dexline").AtRect(206, 244, 380, 16).BindText "dexLine"

    ui.Label("statshdr").AtRect(24, 314, 200, 16).Text("Base stats").Bold
    For statRow = 1 To 6
        ui.Label("statn" & statRow) _
            .AtRect(24, 320 + statRow * 22, 64, 16).Text(StatCaption(statRow))
        ui.ProgressBar("statb" & statRow) _
            .AtRect(94, 323 + statRow * 22, 220, 10).BindValue StatPctKey(statRow)
        ui.Label("statv" & statRow) _
            .AtRect(322, 320 + statRow * 22, 44, 16).BindText StatKey(statRow)
    Next statRow

    ui.Button("catch").AtRect(400, 342, 202, 34).Text("Catch!").Primary _
        .OnClick "PokeDex.CatchCurrent"
    ui.Button("hunt").AtRect(400, 386, 130, 26).Text("Shiny hunt") _
        .Secondary.OnClick "PokeDex.StartShinyHunt"
    ui.Button("stophunt").AtRect(538, 386, 64, 26).Text("Stop").Danger _
        .OnClick "PokeDex.StopShinyHunt"
    ui.Button("hunt").BindEnabled "hunting", True
    ui.Button("stophunt").BindEnabled "hunting"
    ui.ProgressBar("huntbar").AtRect(400, 420, 202, 8).BindValue "huntPct"

    ui.SetStateDefault "fetchStatus", "idle"
    ui.SetStateDefault "pokeName", "Loading the Kanto dex..."
    ui.SetStateDefault "typeAOn", False
    ui.SetStateDefault "typeBOn", False
    ui.SetStateDefault "sizeLine", ""
    ui.SetStateDefault "dexLine", ""
    ui.SetStateDefault "spriteFile", ""
    ui.SetStateDefault "hunting", False
    ui.SetStateDefault "huntPct", 0
    ui.SetStateDefault "tweening", False
    ui.SetStateDefault "namesLoaded", False
    ui.SetStateDefault "dexId", 0
    ui.HotKey "{F2}", "PokeDex.GoPrev"
    ui.HotKey "{F3}", "PokeDex.GoNext"
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub BuildTeam()
    Dim ui As ReDimUI
    Dim host As Worksheet

    Set host = EnsureSheet("DexTeam")
    Set ui = ReDimUI.Mount(host, "dexteam")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.SetTheme PokeTheme(False)
    ui.AsWindow.WindowTitle "Team"
    ui.NavBar
    ui.OnShow "PokeDex.HandleTeamShown"

    ui.Label("title").AtRect(24, 46, 300, 26).Text("Team builder") _
        .FontSize(18).Bold
    ui.Label("hint").AtRect(24, 76, 420, 16) _
        .Text("Catch species in the Pokedex, then build a party of six.")
    ui.TransferList("team").AtRect 24, 100, 420, 152
    ui.TransferList("team").Captions("Caught", "Party") _
        .WritesTo("party").OnChange "PokeDex.HandlePartyChange"

    ui.Label("preflbl").AtRect(470, 76, 150, 16).Text("House rules").Bold
    ui.CheckList("prefs").AtRect 470, 100, 160, 152
    ui.CheckList("prefs").ItemsFrom( _
        Array("Nicknames", "Auto-heal", "Hard mode", "Shiny only")) _
        .CheckedFrom(Array("Auto-heal")) _
        .WritesTo "houseRules"

    ui.Label("noteslbl").AtRect(24, 268, 200, 16).Text("Strategy notes").Bold
    ui.TextInput("strategy").AtRect(24, 288, 300, 58).MultiLine _
        .WritesTo "strategy"
    ui.Label("partylbl").AtRect(340, 288, 290, 58) _
        .BindText "party", "Party: {0}"
    ui.Label("visits").AtRect(24, 360, 300, 16).BindText _
        "teamVisits", "Window shown {0} times this session."

    ui.SetStateDefault "party", ""
    ui.SetStateDefault "strategy", ""
    ui.SetStateDefault "teamVisits", 0
    ui.Render
    ui.ProtectSurface
End Sub

Private Sub BuildTrainer()
    Dim ui As ReDimUI
    Dim host As Worksheet

    Set host = EnsureSheet("DexTrainer")
    Set ui = ReDimUI.Mount(host, "dextrainer")
    ui.ProtectSurface False
    ui.PrepareCanvas
    ui.SetTheme PokeTheme(False)
    ui.AsWindow.WindowTitle "Trainer"
    ui.NavBar
    ui.OnShow "PokeDex.RefreshTrainerCard"

    ui.Label("title").AtRect(24, 46, 300, 26).Text("Trainer card") _
        .FontSize(18).Bold
    ui.Label("namelbl").AtRect(24, 84, 120, 18).Text("Name")
    ui.TextInput("trainer").AtRect(120, 82, 170, 22).WritesTo "trainerName"
    ui.Label("starterlbl").AtRect(24, 118, 120, 18).Text("Starter").Bold
    ui.RadioGroup("starter").AtRect(24, 140, 160, 62) _
        .Items("Bulbasaur", "Charmander", "Squirtle").Value(1) _
        .WritesTo("starter").OnChange "PokeDex.HandleStarterPick"
    ui.Label("favlbl").AtRect(220, 118, 120, 18).Text("Favorite type").Bold
    ui.SelectBox("favtype").AtRect(220, 140, 140, 24) _
        .Items("Fire", "Water", "Grass", "Electric", "Dragon") _
        .Value(2).WritesTo "favType"
    ui.Label("lvllbl").AtRect(220, 178, 120, 18).Text("Ambition (level)")
    ui.Stepper("level").AtRect(220, 198, 140, 24).SliderRange(1, 100, 5) _
        .Value(5).WritesTo "ambition"

    ui.Toggle("darkmode").AtRect(24, 226, 44, 22).Text("Night mode").WritesTo("darkMode") _
        .OnChange "PokeDex.ApplyThemeChoice"

    ui.Card("card").AtRect(24, 262, 420, 96).Text("Trainer summary")
    ui.Label("cardbody").AtRect(36, 290, 396, 60).BindText "cardText"
    ui.Button("reset").AtRect(470, 262, 150, 30).Text("Reset journey") _
        .Danger.OnClick "PokeDex.ConfirmReset"

    ui.SetStateDefault "trainerName", "Red"
    ui.SetStateDefault "starter", "Bulbasaur"
    ui.SetStateDefault "favType", "Water"
    ui.SetStateDefault "ambition", 5
    ui.SetStateDefault "darkMode", False
    ui.SetStateDefault "cardText", ""
    ui.OnStateChanged Array("trainerName", "starter", "favType", "ambition"), _
        "PokeDex.RefreshTrainerCard"
    ui.Render
    ui.ProtectSurface
End Sub

' =====================================================================
' Fetching: names list once, then details per selection
' =====================================================================

Public Sub HandleBrowseShown()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    If CBool(ui.StateOrDefault("namesLoaded", False)) Then Exit Sub
    If ui.Async("names").IsRunning Then Exit Sub
    Set gNamesTask = ROneCOne.HttpClient.GetStringAsync( _
        API_BASE & "?limit=" & DEX_MAX)
    With ui.Async("names")
        .RunsTask gNamesTask
        .ShowsSpinner "spn"
        .TracksState "fetchStatus"
        .OnDone "PokeDex.ApplyNamesList"
        .OnFail "PokeDex.NamesFailed"
    End With
    ui.Async("names").Start
End Sub

Public Sub ApplyNamesList()
    Dim ui As ReDimUI
    Dim jsonDoc As ROneCOne
    Dim resultsList As ROneCOne
    Dim entry As ROneCOne
    Dim speciesNames As Collection
    Dim idx As Long

    Set ui = ReDimUI.App("dexbrowse")
    Set jsonDoc = ROneCOne.Json.Deserialize(CStr(gNamesTask.Result))
    Set resultsList = jsonDoc.Item("results")
    Set speciesNames = New Collection
    For idx = 0 To resultsList.Count - 1
        Set entry = resultsList.Item(idx)
        speciesNames.Add StrConv(CStr(entry.Item("name")), vbProperCase)
    Next idx
    ui.ComboBox("species").ItemsFrom speciesNames
    ui.SetState "namesLoaded", True
    ui.Toast speciesNames.Count & " Kanto species loaded. Pick one!", 3500
    FetchByPath "1"
End Sub

Public Sub NamesFailed()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    ui.SetState "pokeName", "PokeAPI unreachable."
    ui.Toast "Could not reach PokeAPI: " & ui.AsyncError("names"), 6000
End Sub

Public Sub HandleSpeciesPick()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    FetchByPath LCase$(CStr(ui.StateOrDefault("speciesPick", "")))
End Sub

Public Sub GoPrev()
    StepDex -1
End Sub

Public Sub GoNext()
    StepDex 1
End Sub

Private Sub StepDex(ByVal stepBy As Long)
    Dim ui As ReDimUI
    Dim dexId As Long

    Set ui = ReDimUI.App("dexbrowse")
    dexId = gNavId
    If dexId < 1 Then dexId = CLng(ui.StateOrDefault("dexId", 0))
    dexId = dexId + stepBy
    If dexId < 1 Then dexId = DEX_MAX
    If dexId > DEX_MAX Then dexId = 1
    gNavId = dexId
    FetchByPath CStr(dexId)
End Sub

' Latest wins: a cached payload applies instantly, and a busy op just
' records the newest intent, which the op chases when it settles.
Private Sub FetchByPath(ByVal pathPart As String)
    Dim ui As ReDimUI
    Dim cached As ROneCOne

    If LenB(pathPart) = 0 Then Exit Sub
    pathPart = LCase$(pathPart)
    gWantPath = pathPart
    If pathPart = gShownPath Then Exit Sub
    If TryCachedDetail(pathPart, cached) Then
        gShownPath = pathPart
        ApplyDetailDoc cached
        Exit Sub
    End If
    Set ui = ReDimUI.App("dexbrowse")
    If ui.Async("detail").IsRunning Then Exit Sub
    gFlightPath = pathPart
    Set gDetailTask = ROneCOne.HttpClient.GetStringAsync(API_BASE & pathPart)
    With ui.Async("detail")
        .RunsTask gDetailTask
        .ShowsSpinner "spn"
        .TracksState "fetchStatus"
        .OnDone "PokeDex.ApplyDetails"
        .OnFail "PokeDex.DetailFailed"
    End With
    ui.Async("detail").Start
End Sub

Public Sub ApplyDetails()
    Dim jsonDoc As ROneCOne

    ' A PokeAPI response runs a few hundred KB and the dex draws under
    ' 2 KB of it. The allowlist materializes those members and steps
    ' over the rest, so the move lists and per-game sprite variants are
    ' never built at all.
    Set jsonDoc = ROneCOne.Json.DeserializeOnly( _
        CStr(gDetailTask.Result), DetailPaths())
    CacheDetail gFlightPath, jsonDoc
    If gFlightPath = gWantPath Then
        gShownPath = gFlightPath
        ApplyDetailDoc jsonDoc
    ElseIf gWantPath <> gShownPath Then
        ' The user moved on mid-flight; chase the newest intent.
        FetchByPath gWantPath
    End If
End Sub

' Everything the browse window draws, and nothing else.
Private Function DetailPaths() As Variant
    DetailPaths = Array("$.id", "$.name", "$.height", "$.weight", _
        "$.types", "$.stats", "$.sprites.front_default")
End Function

Private Sub ApplyDetailDoc(ByVal jsonDoc As ROneCOne)
    Dim ui As ReDimUI
    Dim typesList As ROneCOne
    Dim statsList As ROneCOne
    Dim entry As ROneCOne
    Dim inner As ROneCOne
    Dim speciesName As String
    Dim dexId As Long
    Dim idx As Long
    Dim statIndex As Long

    Set ui = ReDimUI.App("dexbrowse")
    dexId = CLng(jsonDoc.Item("id"))
    speciesName = StrConv(CStr(jsonDoc.Item("name")), vbProperCase)
    CacheDetail CStr(dexId), jsonDoc
    CacheDetail LCase$(CStr(jsonDoc.Item("name"))), jsonDoc
    gNavId = dexId

    ui.BeginUpdate
    ui.SetState "dexId", dexId
    ui.SetState "pokeName", "#" & Format$(dexId, "000") & "  " & speciesName
    ui.SetState "sizeLine", _
        Format$(CDbl(jsonDoc.Item("height")) / 10, "0.0") & " m  /  " & _
        Format$(CDbl(jsonDoc.Item("weight")) / 10, "0.0") & " kg"
    ui.SetState "dexLine", "Kanto dex " & dexId & " of " & DEX_MAX

    Set typesList = jsonDoc.Item("types")
    ApplyTypeBadge ui, "typea", "typeA", typesList, 0
    ApplyTypeBadge ui, "typeb", "typeB", typesList, 1

    Set statsList = jsonDoc.Item("stats")
    For idx = 0 To statsList.Count - 1
        Set entry = statsList.Item(idx)
        Set inner = entry.Item("stat")
        statIndex = StatIndexOf(CStr(inner.Item("name")))
        If statIndex > 0 Then
            gStatTarget(statIndex) = CDbl(entry.Item("base_stat"))
        End If
    Next idx
    ui.EndUpdate
    StartStatTween

    ui.ComboBox("species").InputValue = speciesName
    ApplySprite ui, jsonDoc, dexId
End Sub

Public Sub DetailFailed()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    ui.Toast "Fetch failed: " & ui.AsyncError("detail"), 5000
    If gWantPath <> gFlightPath And gWantPath <> gShownPath Then
        FetchByPath gWantPath
    End If
End Sub

Private Function TryCachedDetail( _
    ByVal cacheKey As String, _
    ByRef jsonDoc As ROneCOne _
) As Boolean
    Dim cached As ROneCOne

    EnsureDetailCache
    If Not gDetailCache.ContainsKey(cacheKey) Then Exit Function
    Set cached = gDetailCache.Item(cacheKey)
    Set jsonDoc = cached
    TryCachedDetail = True
End Function

Private Sub CacheDetail(ByVal cacheKey As String, ByVal jsonDoc As ROneCOne)
    If LenB(cacheKey) = 0 Then Exit Sub
    EnsureDetailCache
    If gDetailCache.ContainsKey(cacheKey) Then Exit Sub
    gDetailCache.Add cacheKey, jsonDoc
End Sub

' A typed dictionary rather than a bare Collection: real string keys,
' membership asked directly instead of trapped from a failed lookup.
Private Sub EnsureDetailCache()
    If gDetailCache Is Nothing Then
        Set gDetailCache = ROneCOne.DictionaryOf(vbString, vbObject)
    End If
End Sub

Private Sub ApplyTypeBadge( _
    ByVal ui As ReDimUI, _
    ByVal badgeId As String, _
    ByVal stateKey As String, _
    ByVal typesList As ROneCOne, _
    ByVal slotIndex As Long _
)
    Dim entry As ROneCOne
    Dim inner As ROneCOne
    Dim typeLabel As String

    If slotIndex < typesList.Count Then
        Set entry = typesList.Item(slotIndex)
        Set inner = entry.Item("type")
        typeLabel = CStr(inner.Item("name"))
        ui.Label(badgeId).Fill(TypeColor(typeLabel)) _
            .TextColor RGB(255, 255, 255)
        ui.SetState stateKey, "  " & StrConv(typeLabel, vbProperCase)
        ui.SetState stateKey & "On", True
    Else
        ui.SetState stateKey & "On", False
    End If
End Sub

Private Sub ApplySprite( _
    ByVal ui As ReDimUI, _
    ByVal jsonDoc As ROneCOne, _
    ByVal dexId As Long _
)
    Dim sprites As ROneCOne
    Dim spriteUrl As Variant
    Dim localPath As String

    ' An allowlist omits a path the response did not carry, so the
    ' members are checked rather than assumed.
    If Not jsonDoc.ContainsKey("sprites") Then Exit Sub
    Set sprites = jsonDoc.Item("sprites")
    If Not sprites.ContainsKey("front_default") Then Exit Sub
    spriteUrl = sprites.Item("front_default")
    If IsNull(spriteUrl) Or LenB(CStr(spriteUrl)) = 0 Then Exit Sub
    localPath = SpritePath(dexId)
    If ROneCOne.File.Exists(localPath) Then
        ui.SetState "spriteFile", localPath
        Exit Sub
    End If
    ' Not on disk yet, so fetch it through the pump instead of blocking
    ' the frame. The previous sprite stays up until the new one lands,
    ' which reads better than flashing a placeholder.
    gSpriteWantId = dexId
    gSpriteWantUrl = CStr(spriteUrl)
    StartSpriteFetch ui
End Sub

Private Sub StartSpriteFetch(ByVal ui As ReDimUI)
    If gSpriteWantId = 0 Then Exit Sub
    If ui.Async("sprite").IsRunning Then Exit Sub
    gSpriteFlightId = gSpriteWantId
    Set gSpriteTask = ROneCOne.HttpClient.DownloadFileAsync( _
        gSpriteWantUrl, SpritePath(gSpriteFlightId))
    With ui.Async("sprite")
        .RunsTask gSpriteTask
        .ShowsSpinner "spn"
        .OnDone "PokeDex.SpriteReady"
        .OnFail "PokeDex.SpriteFailed"
    End With
    ui.Async("sprite").Start
End Sub

Public Sub SpriteReady()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    ' Paint it only if the dex still shows the species it belongs to; a
    ' fast switch can land a sprite nobody is looking at any more.
    If gSpriteFlightId = CLng(ui.StateOrDefault("dexId", 0)) Then
        ui.SetState "spriteFile", SpritePath(gSpriteFlightId)
    End If
    If gSpriteWantId <> gSpriteFlightId Then StartSpriteFetch ui
End Sub

Public Sub SpriteFailed()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    ui.Toast "Sprite download failed: " & ui.AsyncError("sprite"), 4000
    If gSpriteWantId <> gSpriteFlightId Then StartSpriteFetch ui
End Sub

Private Function SpritePath(ByVal dexId As Long) As String
    SpritePath = ROneCOne.Path.Combine( _
        ROneCOne.Path.GetTempPath(), "redex_" & dexId & ".png")
End Function

' =====================================================================
' Stat tween: bars ease toward their targets on a 16 ms paced job
' =====================================================================

Private Sub StartStatTween()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    If CBool(ui.StateOrDefault("tweening", False)) Then Exit Sub
    ui.SetState "tweening", True
    ui.Job("tween").Steps("PokeDex.TweenStats").PacedMs 33
    ui.Job("tween").StartJob
End Sub

Public Function TweenStats() As Boolean
    Dim ui As ReDimUI
    Dim statIndex As Long
    Dim allDone As Boolean

    Set ui = ReDimUI.App("dexbrowse")
    allDone = True
    ui.BeginUpdate
    For statIndex = 1 To 6
        gStatNow(statIndex) = gStatNow(statIndex) + _
            (gStatTarget(statIndex) - gStatNow(statIndex)) * 0.35
        If Abs(gStatTarget(statIndex) - gStatNow(statIndex)) > 0.8 Then
            allDone = False
        Else
            gStatNow(statIndex) = gStatTarget(statIndex)
        End If
        ui.SetState StatKey(statIndex), CLng(gStatNow(statIndex))
        ui.SetState StatPctKey(statIndex), _
            CLng(gStatNow(statIndex) / 1.6)
    Next statIndex
    ui.EndUpdate
    If allDone Then ui.SetState "tweening", False
    TweenStats = allDone
End Function

' =====================================================================
' Catching, shiny hunting, party rules
' =====================================================================

Public Sub CatchCurrent()
    Dim browseApp As ReDimUI
    Dim teamApp As ReDimUI
    Dim caughtName As String

    Set browseApp = ReDimUI.App("dexbrowse")
    caughtName = CStr(browseApp.StateOrDefault("pokeName", ""))
    If InStr(caughtName, "  ") = 0 Then Exit Sub
    caughtName = Mid$(caughtName, InStr(caughtName, "  ") + 2)
    Set teamApp = ReDimUI.App("dexteam")
    If teamApp.TransferList("team").ItemPosition(caughtName) > 0 Then
        browseApp.Toast caughtName & " is already in your box.", 3000
        Exit Sub
    End If
    ' The party is the chosen side; a Pokemon moved there is caught too.
    If teamApp.TransferList("team").ChosenPosition(caughtName) > 0 Then
        browseApp.Toast caughtName & " is already in your party.", 3000
        Exit Sub
    End If
    teamApp.TransferList("team").AddItem caughtName
    browseApp.Toast "Gotcha! " & caughtName & " was caught!", 3500
End Sub

Public Sub StartShinyHunt()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    If CBool(ui.StateOrDefault("hunting", False)) Then Exit Sub
    gHuntStep = 0
    ui.SetState "hunting", True
    ui.SetState "huntPct", 0
    ui.Job("hunt").Steps("PokeDex.HuntStep").PacedMs 60
    ui.Job("hunt").JobOnDone "PokeDex.HuntDone"
    ui.Job("hunt").StartJob
End Sub

Public Function HuntStep() As Boolean
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    gHuntStep = gHuntStep + 1
    ui.SetState "huntPct", gHuntStep
    HuntStep = (gHuntStep >= 100)
End Function

Public Sub HuntDone()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    ui.SetState "hunting", False
    ui.SetState "huntPct", 0
    If CLng(ui.StateOrDefault("dexId", 0)) Mod 8 = 1 Then
        ui.Toast "It sparkles... a SHINY appeared!", 6000
    Else
        ui.Toast "No shiny this time. The hunt continues.", 4000
    End If
End Sub

Public Sub StopShinyHunt()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexbrowse")
    If Not CBool(ui.StateOrDefault("hunting", False)) Then Exit Sub
    ui.CancelJob "hunt"
    ui.SetState "hunting", False
    ui.SetState "huntPct", 0
    ui.Toast "Hunt called off.", 3000
End Sub

Public Sub HandlePartyChange()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexteam")
    If ui.TransferList("team").ChosenCount > 6 Then
        ui.Toast "A party carries six! The rest ride in the box.", 4000
    End If
End Sub

Public Sub HandleTeamShown()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dexteam")
    ui.SetState "teamVisits", CLng(ui.StateOrDefault("teamVisits", 0)) + 1
End Sub

Public Sub HandleStarterPick()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dextrainer")
    ui.Toast CStr(ui.State("starter")) & ", I choose you!", 3000
    FetchByPath LCase$(CStr(ui.State("starter")))
    ReDimUI.Navigate "dexbrowse"
End Sub

' =====================================================================
' Trainer card, theming, reset
' =====================================================================

Public Sub RefreshTrainerCard()
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dextrainer")
    ui.SetState "cardText", _
        "Trainer " & CStr(ui.StateOrDefault("trainerName", "Red")) & _
        "   Starter: " & CStr(ui.StateOrDefault("starter", "-")) & vbLf & _
        "Favorite type: " & CStr(ui.StateOrDefault("favType", "-")) & _
        "   Aiming for level " & CStr(ui.StateOrDefault("ambition", 5))
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
    Dim ui As ReDimUI

    Set ui = ReDimUI.App("dextrainer")
    ui.Confirm "Reset journey?", _
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

' A Pokedex-red theme in light and night variants: a stock theme with
' every color replaced through the theme builders.
Private Function PokeTheme(ByVal darkMode As Boolean) As ReDimUI
    Dim themeValue As ReDimUI

    If darkMode Then
        Set themeValue = ReDimUI.ThemeDark _
            .WithPrimary(RGB(232, 84, 74), RGB(20, 20, 22)) _
            .WithSurface(RGB(45, 44, 48), RGB(240, 238, 235)) _
            .WithMuted(RGB(66, 64, 70), RGB(176, 172, 168)) _
            .WithStatus(RGB(96, 200, 140), RGB(240, 110, 110)) _
            .WithBorder(RGB(92, 88, 94)).WithCanvas(RGB(30, 29, 33)) _
            .WithFont("Segoe UI", 11)
    Else
        Set themeValue = ReDimUI.ThemeLight _
            .WithPrimary(RGB(214, 55, 46), RGB(255, 255, 255)) _
            .WithSurface(RGB(255, 255, 255), RGB(40, 40, 42)) _
            .WithMuted(RGB(238, 233, 229), RGB(122, 116, 112)) _
            .WithStatus(RGB(46, 140, 90), RGB(178, 34, 52)) _
            .WithBorder(RGB(220, 212, 206)).WithCanvas(RGB(250, 247, 244)) _
            .WithFont("Segoe UI", 11)
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
Private Function TypeColor(ByVal typeLabel As String) As Long
    Select Case typeLabel
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
