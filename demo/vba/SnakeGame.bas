Attribute VB_Name = "SnakeGame"
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

' Snake: proof that the ReDim pump is a real render loop. The chrome (score,
' buttons, pause toggle, game-over modal) is ReDim components; the board is
' painted straight onto cells; movement runs as a pump-driven job at roughly
' twenty ticks per second with arrow-key steering via Application.OnKey.

Private Const APP_ID As String = "snake"
Private Const BOARD_TOP As Long = 5
Private Const BOARD_LEFT As Long = 8
Private Const BOARD_SIZE As Long = 18

Private gBody As Collection
Private gDirRow As Long
Private gDirCol As Long
Private gNextDirRow As Long
Private gNextDirCol As Long
Private gFoodRow As Long
Private gFoodCol As Long
Private gScore As Long
Private gGameActive As Boolean

Public Sub Auto_Open()
    BuildSnake
End Sub

Public Function SnakeApp() As ReDimUI
    Set SnakeApp = ReDimUI.App(APP_ID)
End Function

Public Sub BuildSnake()
    Dim ui As ReDimUI
    Dim host As Worksheet
    Dim boardCells As Range

    Set host = ThisWorkbook.Worksheets(1)
    Set ui = ReDimUI.Mount(host, APP_ID)
    ui.ProtectSurface False
    ui.PrepareCanvas

    Set boardCells = BoardArea(host)
    boardCells.ColumnWidth = 2.2
    host.Rows(BOARD_TOP & ":" & BOARD_TOP + BOARD_SIZE - 1).RowHeight = 14

    ui.Label("title").AtRect(24, 12, 260, 28).Text("ReDim Snake") _
        .FontSize(18).Bold
    ui.Label("score").AtRect(24, 44, 200, 20) _
        .BindText("score", "Score: {0}").FontSize(12).Bold
    ui.Button("newgame").AtRect(24, 76, 110, 30).Text("New game").Primary _
        .OnClick "SnakeGame.NewGame"
    ui.Toggle("pause").AtRect(144, 80, 44, 22).WritesTo "paused"
    ui.Label("pauselbl").AtRect(196, 82, 60, 18).Text("Pause")
    ui.Label("hint").AtRect(24, 116, 220, 60) _
        .Text("Steer with the arrow keys. Eat the red squares. Do not eat yourself.")

    ui.SetState "score", 0
    ui.SetState "paused", False
    ui.Render
    PaintBoard
    HookKeys
    ui.ProtectSurface
End Sub

Private Function BoardArea(ByVal host As Worksheet) As Range
    Set BoardArea = host.Range( _
        host.Cells(BOARD_TOP, BOARD_LEFT), _
        host.Cells(BOARD_TOP + BOARD_SIZE - 1, BOARD_LEFT + BOARD_SIZE - 1))
End Function

Private Sub PaintBoard()
    Dim ui As ReDimUI
    Dim area As Range

    Set ui = SnakeApp()
    Set area = BoardArea(ui.Sheet)
    area.Interior.Color = ui.Theme.SurfaceColor
    area.BorderAround Color:=ui.Theme.BorderColor, Weight:=xlMedium
End Sub

Public Sub HookKeys()
    With SnakeApp()
        .HotKey "{UP}", "SnakeGame.KeyUp"
        .HotKey "{DOWN}", "SnakeGame.KeyDown"
        .HotKey "{LEFT}", "SnakeGame.KeyLeft"
        .HotKey "{RIGHT}", "SnakeGame.KeyRight"
    End With
End Sub

Public Sub UnhookKeys()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
End Sub

Public Sub KeyUp()
    SetDirection -1, 0
End Sub

Public Sub KeyDown()
    SetDirection 1, 0
End Sub

Public Sub KeyLeft()
    SetDirection 0, -1
End Sub

Public Sub KeyRight()
    SetDirection 0, 1
End Sub

' Reversal into your own neck is ignored; every other turn is queued for the
' next movement tick so fast key taps cannot double-turn inside one move.
Private Sub SetDirection(ByVal dirRow As Long, ByVal dirCol As Long)
    If Not gGameActive Then Exit Sub
    If dirRow = -gDirRow And dirCol = -gDirCol Then Exit Sub
    gNextDirRow = dirRow
    gNextDirCol = dirCol
End Sub

Public Sub NewGame()
    Dim ui As ReDimUI
    Dim originRow As Long
    Dim originCol As Long
    Dim segment As Long

    Set ui = SnakeApp()
    ui.CloseModal
    PaintBoard
    Set gBody = New Collection
    originRow = BOARD_TOP + BOARD_SIZE \ 2
    originCol = BOARD_LEFT + BOARD_SIZE \ 2 - 2
    For segment = 2 To 0 Step -1
        gBody.Add EncodeCell(originRow, originCol - segment)
    Next segment
    gDirRow = 0: gDirCol = 1
    gNextDirRow = 0: gNextDirCol = 1
    gScore = 0
    gGameActive = True
    ui.SetState "score", 0
    ui.SetState "paused", False
    PaintSnakeFull
    PlaceFood
    HookKeys
    ui.Job("loop").Steps("SnakeGame.GameStep").PacedMs(150) _
        .JobOnDone "SnakeGame.GameOver"
    If Not ui.Job("loop").JobIsRunning Then ui.Job("loop").StartJob
End Sub

' One paced step is one movement frame. Returning True ends the job, which
' is game over.
Public Function GameStep() As Boolean
    Dim ui As ReDimUI

    On Error GoTo Dead
    Set ui = SnakeApp()
    If Not gGameActive Then
        GameStep = True
        Exit Function
    End If
    If CBool(ui.StateOrDefault("paused", False)) Then Exit Function
    GameStep = Not MoveSnake(ui)
    Exit Function

Dead:
    GameStep = True
End Function

Private Function MoveSnake(ByVal ui As ReDimUI) As Boolean
    Dim headRow As Long
    Dim headCol As Long
    Dim nextRow As Long
    Dim nextCol As Long
    Dim ate As Boolean

    gDirRow = gNextDirRow
    gDirCol = gNextDirCol
    DecodeCell CLng(gBody.Item(gBody.Count)), headRow, headCol
    nextRow = headRow + gDirRow
    nextCol = headCol + gDirCol

    If nextRow < BOARD_TOP Or nextRow > BOARD_TOP + BOARD_SIZE - 1 _
        Or nextCol < BOARD_LEFT Or nextCol > BOARD_LEFT + BOARD_SIZE - 1 Then
        MoveSnake = False
        Exit Function
    End If
    If BodyContains(nextRow, nextCol) Then
        MoveSnake = False
        Exit Function
    End If

    ate = (nextRow = gFoodRow And nextCol = gFoodCol)
    gBody.Add EncodeCell(nextRow, nextCol)
    ui.Sheet.Cells(nextRow, nextCol).Interior.Color = ui.Theme.PrimaryColor
    If ate Then
        gScore = gScore + 1
        ui.SetState "score", gScore
        ' Speed up as the snake grows, floored at a playable pace.
        Dim pace As Long
        pace = 150 - gScore * 4
        If pace < 80 Then pace = 80
        ui.Job("loop").PacedMs pace
        PlaceFood
    Else
        Dim tailRow As Long
        Dim tailCol As Long
        DecodeCell CLng(gBody.Item(1)), tailRow, tailCol
        gBody.Remove 1
        ui.Sheet.Cells(tailRow, tailCol).Interior.Color = ui.Theme.SurfaceColor
    End If
    MoveSnake = True
End Function

Public Sub GameOver()
    Dim ui As ReDimUI

    gGameActive = False
    Set ui = SnakeApp()
    ui.Toast "Game over.", 2000
    ui.Confirm "Game over", "Final score: " & gScore & ". Play again?", _
        "SnakeGame.NewGame", vbNullString, "Play again", "Done"
End Sub

Private Sub PaintSnakeFull()
    Dim ui As ReDimUI
    Dim entry As Variant
    Dim segRow As Long
    Dim segCol As Long

    Set ui = SnakeApp()
    For Each entry In gBody
        DecodeCell CLng(entry), segRow, segCol
        ui.Sheet.Cells(segRow, segCol).Interior.Color = ui.Theme.PrimaryColor
    Next entry
End Sub

Private Sub PlaceFood()
    Dim ui As ReDimUI
    Dim tryRow As Long
    Dim tryCol As Long

    Set ui = SnakeApp()
    Randomize
    Do
        tryRow = BOARD_TOP + Int(Rnd * BOARD_SIZE)
        tryCol = BOARD_LEFT + Int(Rnd * BOARD_SIZE)
    Loop While BodyContains(tryRow, tryCol)
    gFoodRow = tryRow
    gFoodCol = tryCol
    ui.Sheet.Cells(tryRow, tryCol).Interior.Color = ui.Theme.DangerColor
End Sub

Private Function BodyContains(ByVal cellRow As Long, ByVal cellCol As Long) As Boolean
    Dim entry As Variant

    If gBody Is Nothing Then Exit Function
    For Each entry In gBody
        If CLng(entry) = EncodeCell(cellRow, cellCol) Then
            BodyContains = True
            Exit Function
        End If
    Next entry
End Function

Private Function EncodeCell(ByVal cellRow As Long, ByVal cellCol As Long) As Long
    EncodeCell = cellRow * 1000 + cellCol
End Function

Private Sub DecodeCell( _
    ByVal encoded As Long, _
    ByRef cellRow As Long, _
    ByRef cellCol As Long _
)
    cellRow = encoded \ 1000
    cellCol = encoded Mod 1000
End Sub

' Test seam: expose the head position without touching game internals.
Public Function HeadPosition() As String
    Dim headRow As Long
    Dim headCol As Long

    If gBody Is Nothing Then
        HeadPosition = "none"
    Else
        DecodeCell CLng(gBody.Item(gBody.Count)), headRow, headCol
        HeadPosition = headRow & "," & headCol
    End If
End Function

Public Function IsGameActive() As Boolean
    IsGameActive = gGameActive
End Function
