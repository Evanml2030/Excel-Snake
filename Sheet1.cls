VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Enum Direction
    North = 1
    South = 2
    East = 3
    West = 4
End Enum

Public ws As Worksheet
Public snakeParts As Collection
Public currentRow As Long
Public currentColumn As Long
Public directionSnake As Direction

    Sub RunGame()
        Set ws = ActiveWorkbook.Sheets("Game")
        Set snakeParts = New Collection
        
        Dim gameOver As Boolean
        gameOver = False
        
        Dim TimerGame As TimerWin64
        Set TimerGame = New TimerWin64
        
        Dim TimerBlueSquare As TimerWin64
        Set TimerBlueSquare = New TimerWin64
        
        Dim TimerYellowSquare As TimerWin64
        Set TimerYellowSquare = New TimerWin64
        
        Dim SnakePartNew As snakepart
        Set SnakePartNew = New snakepart
        
        GameBoardReset
        DirectionSnakeInitialize
        StartPositionInitalize
        StartGameBoardInitalize
        TimerGame.TimerSet
        TimerBlueSquare.TimerSet
        TimerYellowSquare.TimerSet
        
        ws.cells(currentRow, currentColumn).Select
        Do While gameOver = False
            If TimerGame.CheckQuarterSecondPassed = True Then
                CurrentCellUpdate
                ws.cells(currentRow, currentColumn).Select
                If SnakePartOverlapItself(currentRow, currentColumn) = True Then
                    gameOver = True
                    Exit Do
                ElseIf SnakePartYellowSquareOverlap = True Then
                    gameOver = True
                    Exit Do
                ElseIf SnakePartBlueSquareOverlap = True Then
                    Call SnakePartAdd(currentRow, currentColumn)
                    Call SnakePartAdd(currentRow, currentColumn)
                    Call SnakePartAdd(currentRow, currentColumn)
                    Call SnakePartRemove
                    ws.cells(currentRow, currentColumn).Select
                    TimerGame.TimerSet
                Else
                    Call SnakePartAdd(currentRow, currentColumn)
                    Call SnakePartRemove
                    ws.cells(currentRow, currentColumn).Select
                    TimerGame.TimerSet
                End If
            End If
            
            If TimerBlueSquare.CheckFiveSecondsPassed = True Then
                BlueSquareAdd
                TimerBlueSquare.TimerSet
            End If
            
            If TimerYellowSquare.CheckFiveSecondsPassed = True Then
                YellowSquareAdd
                TimerYellowSquare.TimerSet
            End If
            gameOver = OutOfBounds
            DoEvents
        Loop
    End Sub

        Private Sub GameBoardReset()
            ws.cells.Interior.Color = RGB(300, 300, 300)
        End Sub

        Private Sub DirectionSnakeInitialize()
            directionSnake = East
        End Sub
        
        Private Sub StartPositionInitalize()
            currentRow = 96
            currentColumn = 64
        End Sub
        
        Private Sub StartGameBoardInitalize()
            Call SnakePartAdd(currentRow, currentColumn - 6)
            Call SnakePartAdd(currentRow, currentColumn - 5)
            Call SnakePartAdd(currentRow, currentColumn - 4)
            Call SnakePartAdd(currentRow, currentColumn - 3)
            Call SnakePartAdd(currentRow, currentColumn - 2)
            Call SnakePartAdd(currentRow, currentColumn - 1)
            Call SnakePartAdd(currentRow, currentColumn)
        End Sub
        
        Private Sub SnakePartAdd(ByVal row As Long, ByVal column As Long)
            Dim SnakePartNew As snakepart
            Set SnakePartNew = New snakepart
            SnakePartNew.PropertiesSet row, column
            SnakePartAddToCollection SnakePartNew
            SnakePartAddToGameBoard SnakePartNew
        End Sub
        
            Private Sub SnakePartAddToCollection(ByRef snakepart As snakepart)
                snakeParts.add snakepart
            End Sub
            
            Private Sub SnakePartAddToGameBoard(ByRef snakepart As snakepart)
                ws.cells(snakepart.row, snakepart.column).Interior.Color = RGB(0, 150, 0)
            End Sub
        
        Private Sub SnakePartRemove()
            SnakePartRemoveFromGameBoard
            SnakePartRemoveFromCollection
        End Sub
        
            Private Sub SnakePartRemoveFromCollection()
                snakeParts.Remove 1
            End Sub
        
            Private Sub SnakePartRemoveFromGameBoard()
                ws.cells(snakeParts.Item(1).row, snakeParts.Item(1).column).Interior.Color = RGB(300, 300, 300)
            End Sub
            
        Private Function OutOfBounds() As Boolean
            If currentRow < 9 Or _
               currentRow > 189 Or _
               currentColumn < 21 Or _
               currentColumn > 108 Then
                OutOfBounds = True
                MsgBox "GameOver"
            Else
                OutOfBounds = False
            End If
        End Function
        
        Private Function SnakePartOverlapItself(ByVal row As Long, ByVal column As Long) As Boolean
            If ws.cells(row, column).Interior.Color = RGB(0, 150, 0) Then
                MsgBox "GameOver"
                SnakePartOverlapItself = True
            Else
                SnakePartOverlapItself = False
            End If
        End Function
        
        Private Sub BlueSquareAdd()
            Dim TopLeftCornerRow As Long
            Dim TopLeftCornerColumn As Long
            
            TopLeftCornerRow = Application.WorksheetFunction.RandBetween(9, 189)
            TopLeftCornerColumn = Application.WorksheetFunction.RandBetween(21, 108)
            
            ws.cells(TopLeftCornerRow, TopLeftCornerColumn).Interior.Color = RGB(0, 0, 150)
            ws.cells(TopLeftCornerRow, TopLeftCornerColumn + 1).Interior.Color = RGB(0, 0, 150)
            ws.cells(TopLeftCornerRow + 1, TopLeftCornerColumn).Interior.Color = RGB(0, 0, 150)
            ws.cells(TopLeftCornerRow + 1, TopLeftCornerColumn + 1).Interior.Color = RGB(0, 0, 150)
        End Sub
        
        Private Function SnakePartBlueSquareOverlap() As Boolean
            If ws.cells(currentRow, currentColumn).Interior.Color = RGB(0, 0, 150) Then
                SnakePartBlueSquareOverlap = True
            Else
                SnakePartBlueSquareOverlap = False
            End If
        End Function
        
        Private Sub YellowSquareAdd()
            Dim TopLeftCornerRow As Long
            Dim TopLeftCornerColumn As Long
            
            TopLeftCornerRow = Application.WorksheetFunction.RandBetween(9, 189)
            TopLeftCornerColumn = Application.WorksheetFunction.RandBetween(21, 108)
            
            ws.cells(TopLeftCornerRow, TopLeftCornerColumn).Interior.Color = RGB(255, 140, 0)
            ws.cells(TopLeftCornerRow, TopLeftCornerColumn + 1).Interior.Color = RGB(255, 140, 0)
            ws.cells(TopLeftCornerRow + 1, TopLeftCornerColumn).Interior.Color = RGB(255, 140, 0)
            ws.cells(TopLeftCornerRow + 1, TopLeftCornerColumn + 1).Interior.Color = RGB(255, 140, 0)
        End Sub
        
        Private Function SnakePartYellowSquareOverlap() As Boolean
            If ws.cells(currentRow, currentColumn).Interior.Color = RGB(255, 140, 0) Then
                MsgBox "GameOver"
                SnakePartYellowSquareOverlap = True
            Else
                SnakePartYellowSquareOverlap = False
            End If
        End Function
        
        Private Sub CurrentCellUpdate()
            Select Case directionSnake
                Case Is = Direction.North
                    currentRow = currentRow - 1
                Case Is = Direction.South
                    currentRow = currentRow + 1
                Case Is = Direction.East
                    currentColumn = currentColumn + 1
                Case Is = Direction.West
                    currentColumn = currentColumn - 1
            End Select
        End Sub
        
        Private Sub SnakeCollectionUpdate(ByRef snakeParts As Collection)
            snakeParts.add currentRow
        End Sub

    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        'rowSwitch
        If directionSnake = East Or directionSnake = West Then
            If Target.column = currentColumn Then
                If Target.row <> currentRow Then
                    If Target.row = currentRow - 1 Then
                        directionSnake = North
                    ElseIf Target.row = currentRow + 1 Then
                        directionSnake = South
                    End If
                End If
            End If
        End If
        
        'columnSwitch
        If directionSnake = North Or directionSnake = South Then
            If Target.row = currentRow Then
                If Target.column <> currentColumn Then
                    If Target.column = currentColumn + 1 Then
                        directionSnake = East
                    ElseIf Target.column = currentColumn - 1 Then
                        directionSnake = West
                    End If
                End If
            End If
        End If
    End Sub
