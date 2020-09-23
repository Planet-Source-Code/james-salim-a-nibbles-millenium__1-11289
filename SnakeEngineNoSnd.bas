Attribute VB_Name = "SnakeEngine"
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Const ArenaRow = 80              'Arena Rows
Public Const ArenaCol = 60              'Arena Collumns
Public Const MaxSnakeLength = 1000      'Maximum Length for Snake
Public Const Arena_None = 0
Public Const Arena_Wall = 1
Public Const Arena_Snake1 = 2
Public Const Arena_Snake2 = 3
Public Const Arena_Number = 4

Public Const Color_None = 9
Public Const Color_Wall = 4
Public Const Color_Snake1 = 14
Public Const Color_Snake2 = 13
Public Const Color_Number = 15

Public StartInterval As Integer
Public GameSpeed As Integer
Public PlayerNum As Integer
Public BoxHeight As Integer
Public BoxWidth As Integer
Public CurLevel As Integer
Public CurNumber As Integer
Public InitialLive As Integer
Public InitialLevel As Integer
Public ScoreBeforeLive As Integer
Public NumberOnScreen As Boolean
'Public CurMidiNo As Integer
'Public DeadWav As String
'Public PointWav As String
'Public LevelWav As String

Type SnakeChar                          'TYPE for Snake, including Position, Direction and etc
    HeadX As Integer
    HeadY As Integer
    DirX As Integer
    DirY As Integer
    CurLength As Integer
    Length As Integer
    Alive As Boolean
    Lives As Integer
    Score As Long
    BodyX(1 To MaxSnakeLength) As Integer
    BodyY(1 To MaxSnakeLength) As Integer
    PlayerColor As Integer
    PlayerName As String
End Type

Public Snake(1 To 2) As SnakeChar
Public Arena(1 To 80, 1 To 60) As Integer

Sub InitGame()
'Objective: Set Lives, Scores, and other variables
For i = 1 To SnakeEngine.PlayerNum
    Snake(i).Lives = InitialLive
    Snake(i).Score = 0
    Snake(i).Alive = True
Next i
SnakeFrm.StartScreen.Visible = False
SnakeFrm.Timer2.Interval = 0
SnakeFrm.Pause_Game.Enabled = True
MsgClose
InitArena
SetLevel InitialLevel
End Sub

Sub SetArena(Row, Col, Mode)
'Objective: Initiate & Draw a mode on the Arena Box
'Input    : * Row
'           * Col
'           * Mode - 0) None
'                    1) Wall
'                    2) Snake 1 Body
'                    3) Snake 2 Body
'Output   : None

Arena(Row, Col) = Mode

'--Drawing Method--
Select Case Mode
    Case Arena_Wall: AssignedColor = Color_Wall
    Case Arena_Snake1: AssignedColor = Color_Snake1
    Case Arena_Snake2: AssignedColor = Color_Snake2
    Case Arena_Number: AssignedColor = Color_Number
    Case Else: AssignedColor = Color_None
End Select
If AssignedColor = Color_Number Then
    SnakeFrm.Number.Caption = CurNumber
    SnakeFrm.Number.Top = (Col - 1) * BoxHeight - 50
    SnakeFrm.Number.Left = (Row - 1) * BoxWidth + 10
Else
    If SnakeFrm.WireFrame.Checked Then
        SnakeFrm.PlayScreen.Line ((Row - 1) * BoxWidth, (Col - 1) * BoxHeight)-(Row * BoxWidth - 15, Col * BoxHeight - 15), QBColor(AssignedColor), B
    Else
        SnakeFrm.PlayScreen.Line ((Row - 1) * BoxWidth, (Col - 1) * BoxHeight)-(Row * BoxWidth - 15, Col * BoxHeight - 15), QBColor(AssignedColor), BF
    End If
End If
'--End Drawing Method--
Exit Sub
End Sub

Sub InitArena()
'Objective: * To initiate arena background and borders
'           * Initiate any other variable
'Input    : None
'Output   : None

SnakeFrm.WireFrame.Enabled = False
'--Draw Background--
For Col = 1 To ArenaCol
    For Row = 1 To ArenaRow
        SetArena Row, Col, Arena_None
    Next Row
Next Col
'==Draw Background==

'--Draw Border--
For Col = 1 To ArenaCol
    SetArena 1, Col, 1
    SetArena ArenaRow, Col, Arena_Wall
Next Col

For Row = 1 To ArenaRow
    SetArena Row, 1, 1
    SetArena Row, ArenaCol, Arena_Wall
Next Row
'==Draw Border==
For i = 1 To PlayerNum
    For SnakeLen = 1 To Snake(1).CurLength
        Snake(i).BodyX(SnakeLen) = 0
        Snake(i).BodyY(SnakeLen) = 0
    Next SnakeLen
Next i
NumberOnScreen = False
End Sub

Sub SetLevel(Level)
    CurLevel = Level
    Snake(1).Length = 2     'Initialize Snake for level start
    Snake(1).CurLength = 1
    If Snake(1).Lives < 0 Then Snake(1).Alive = False Else Snake(1).Alive = True
    Snake(2).Length = 2
    Snake(2).CurLength = 1
    If Snake(2).Lives < 0 Then Snake(2).Alive = False Else Snake(2).Alive = True
    If Snake(1).Alive Then
        SnakeFrm.StatusBar.Panels.Item(1).Text = Str$(SnakeEngine.Snake(1).Lives) + " Lives"
    Else
        SnakeFrm.StatusBar.Panels.Item(1).Text = ""
    End If
    If Snake(2).Alive Then
        SnakeFrm.StatusBar.Panels.Item(2).Text = Str$(SnakeEngine.Snake(2).Lives) + " Lives"
    Else
        SnakeFrm.StatusBar.Panels.Item(2).Text = ""
    End If
    If Not (Snake(1).Alive And Snake(2).Alive) Then GameOver: Exit Sub
    Snake(1).DirX = 0: Snake(2).DirX = 0: Snake(1).DirY = 0: Snake(2).DirY = 0
    
    Pause
    MsgDraw "Level " + LTrim$(Str$(Level)), "Press F3 to Continue", "Press F2 to start a new game"
    
    Select Case Level
    Case 1
        Snake(1).HeadX = 30: Snake(2).HeadX = 50
        Snake(1).HeadY = 30: Snake(2).HeadY = 30
        Snake(1).DirX = -1: Snake(2).DirX = 1
    Case 2
        For i = 20 To 60
            SetArena i, 30, Arena_Wall
        Next i
        Snake(1).HeadX = 30: Snake(2).HeadX = 50
        Snake(1).HeadY = 40: Snake(2).HeadY = 20
        Snake(1).DirX = -1: Snake(2).DirX = 1
    Case 3
        For i = 10 To 50
            SetArena 20, i, Arena_Wall
            SetArena 60, i, Arena_Wall
        Next i
        Snake(1).HeadX = 25: Snake(2).HeadX = 55
        Snake(1).HeadY = 30: Snake(2).HeadY = 30
        Snake(1).DirY = -1: Snake(2).DirY = 1
    Case 4
        For i = 1 To 30
            SetArena 20, i, Arena_Wall
            SetArena 60, 60 - i, Arena_Wall
        Next i
        For i = 1 To 40
            SetArena i, 40, Arena_Wall
            SetArena 80 - i, 20, Arena_Wall
        Next i
        Snake(1).HeadY = 10: Snake(2).HeadY = 50
        Snake(1).HeadX = 60: Snake(2).HeadX = 20
        Snake(1).DirX = -1: Snake(2).DirX = 1
   
    Case 5
        For i = 13 To 47
            SetArena 20, i, Arena_Wall
            SetArena 60, i, Arena_Wall
        Next i
        For i = 23 To 57
            SetArena i, 10, Arena_Wall
            SetArena i, 50, Arena_Wall
        Next i
        Snake(1).HeadY = 25: Snake(2).HeadY = 25
        Snake(1).HeadX = 50: Snake(2).HeadX = 30
        Snake(1).DirY = -1: Snake(2).DirY = 1

    Case 6
        For i = 2 To 59
            If i < 22 Or i > 38 Then
                SetArena 10, i, Arena_Wall
                SetArena 20, i, Arena_Wall
                SetArena 30, i, Arena_Wall
                SetArena 40, i, Arena_Wall
                SetArena 50, i, Arena_Wall
                SetArena 60, i, Arena_Wall
                SetArena 70, i, Arena_Wall
            End If
        Next i
        Snake(1).HeadY = 10: Snake(2).HeadY = 50
        Snake(1).HeadX = 65: Snake(2).HeadX = 15
        Snake(1).DirY = 1: Snake(2).DirY = -1
   
    Case 7
        For i = 3 To 47
            SetArena i + 2, i + 5, Arena_Wall
            SetArena i + 30, i + 5, Arena_Wall
        Next i
        Snake(1).HeadY = 40: Snake(2).HeadY = 20
        Snake(1).HeadX = 75: Snake(2).HeadX = 5
        Snake(1).DirY = -1: Snake(2).DirY = 1
      
    Case 8
        For i = 2 To 58 Step 2
            SetArena 40, i, Arena_Wall
        Next i
        Snake(1).HeadY = 10: Snake(2).HeadY = 50
        Snake(1).HeadX = 65: Snake(2).HeadX = 15
        Snake(1).DirY = 1: Snake(2).DirY = -1

    Case 9
        For i = 1 To 45
            SetArena 10, i, Arena_Wall
            SetArena 20, 60 - i, Arena_Wall
            SetArena 30, i, Arena_Wall
            SetArena 40, 60 - i, Arena_Wall
            SetArena 50, i, Arena_Wall
            SetArena 60, 60 - i, Arena_Wall
            SetArena 70, i, Arena_Wall
        Next i
        Snake(1).HeadY = 7: Snake(2).HeadY = 53
        Snake(1).HeadX = 65: Snake(2).HeadX = 15
        Snake(1).DirY = 1: Snake(2).DirY = -1

    Case Else
        For i = 2 To 59 Step 2
            SetArena 10, i, Arena_Wall
            SetArena 20, i + 1, Arena_Wall
            SetArena 30, i, Arena_Wall
            SetArena 40, i + 1, Arena_Wall
            SetArena 50, i, Arena_Wall
            SetArena 60, i + 1, Arena_Wall
            SetArena 70, i, Arena_Wall
        Next i
        Snake(1).HeadY = 7: Snake(2).HeadY = 53
        Snake(1).HeadX = 65: Snake(2).HeadX = 15
        Snake(1).DirY = 1: Snake(2).DirY = -1

    End Select
    
    If Snake(1).Alive Then SnakeAddBody Snake(1), Snake(1).HeadX, Snake(1).HeadY, Arena_Snake1
    If PlayerNum > 1 And Snake(2).Alive Then SnakeAddBody Snake(2), Snake(2).HeadX, Snake(2).HeadY, Arena_Snake2
End Sub
Sub SnakeAddBody(Snakes As SnakeChar, Row, Col, ColorMode)
    For i = Snakes.CurLength To 1 Step -1
        Snakes.BodyX(i + 1) = Snakes.BodyX(i)
        Snakes.BodyY(i + 1) = Snakes.BodyY(i)
    Next i
    Snakes.BodyX(1) = Row
    Snakes.BodyY(1) = Col
    Snakes.HeadX = Row
    Snakes.HeadY = Col
    SetArena Row, Col, ColorMode
    
    If Snakes.CurLength < Snakes.Length Then
        Snakes.CurLength = Snakes.CurLength + 1
        Exit Sub
    End If
    
    If Snakes.CurLength = Snakes.Length Then
        If Snakes.BodyX(Snakes.CurLength + 1) <> 0 And Snakes.BodyY(Snakes.CurLength + 1) <> 0 Then
            SetArena Snakes.BodyX(Snakes.CurLength + 1), Snakes.BodyY(Snakes.CurLength + 1), Arena_None
            Snakes.BodyX(Snakes.CurLength + 1) = 0
            Snakes.BodyY(Snakes.CurLength + 1) = 0
        End If
    End If
End Sub

Sub SnakeDead(PlayerNum)
SnakeFrm.Timer1.Interval = 0
Snake(PlayerNum).Alive = False
Snake(PlayerNum).Lives = Snake(PlayerNum).Lives - 1
RestartLevel
End Sub
Sub FinishLevel()
If Snake(1).Alive Then Snake(1).Score = Snake(1).Score + (CurLevel * 100)
If Snake(2).Alive Then Snake(2).Score = Snake(2).Score + (CurLevel * 100)
InitArena
SetLevel CurLevel + 1
CurNumber = 0
End Sub
Sub RestartLevel()
InitArena
SetLevel CurLevel
CurNumber = 0
End Sub
Sub Pause()
If SnakeFrm.Pause_Game.Checked Then
    SnakeFrm.Pause_Game.Checked = False
    MsgClose
    SnakeFrm.Timer1.Interval = GameSpeed
Else
    SnakeFrm.Timer1.Interval = 0
    SnakeFrm.Pause_Game.Checked = True
End If
End Sub
Sub GameOver()
    SnakeFrm.Timer1.Interval = 0
    Pause
    MsgDraw "Game Over", "Press F3 to Continue", "Press F2 to Start a new game"
    SnakeFrm.Pause_Game.Enabled = False
    SnakeFrm.WireFrame.Enabled = True
End Sub
Sub MsgDraw(Message$, F3Message$, F2Message$)
SnakeFrm.PauseScreen.Visible = True
SnakeFrm.PauseCaption = Message$
SnakeFrm.PressF2Label.Caption = F2Message$
SnakeFrm.PressF3Label.Caption = F3Message$
End Sub
Sub MsgClose()
SnakeFrm.PauseScreen.Visible = False
End Sub
