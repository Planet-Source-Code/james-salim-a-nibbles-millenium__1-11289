VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SnakeFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nibbles Millenium - No sound mode"
   ClientHeight    =   6285
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7965
   Icon            =   "SnakeNoSound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox StartScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   0
      Picture         =   "SnakeNoSound.frx":0442
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   531
      TabIndex        =   7
      Top             =   0
      Width           =   8000
      Begin VB.Timer Timer2 
         Left            =   7440
         Top             =   5400
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PlayScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   8000
      Begin VB.PictureBox PauseScreen 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   6000
         Left            =   0
         Picture         =   "SnakeNoSound.frx":45DC
         ScaleHeight     =   5970
         ScaleWidth      =   7965
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   8000
         Begin VB.Label PressF2Label 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3945
            TabIndex        =   6
            Top             =   4200
            Width           =   75
         End
         Begin VB.Label PressF3Label 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3960
            TabIndex        =   5
            Top             =   3840
            Width           =   75
         End
         Begin VB.Label PauseCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1335
            Left            =   0
            TabIndex        =   4
            Top             =   1680
            Width           =   7935
         End
      End
      Begin VB.Timer Timer1 
         Left            =   9600
         Top             =   7320
      End
      Begin VB.Label Number 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   10680
         TabIndex        =   1
         Top             =   8040
         Width           =   75
      End
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Begin VB.Menu P1_Game 
         Caption         =   "&1 Player Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu P2_Game 
         Caption         =   "&2 Player Game"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu Pause_Game 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu Space1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitGame 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "&Option"
      Begin VB.Menu WireFrame 
         Caption         =   "&Wireframe Mode"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Space2 
         Caption         =   "-"
      End
      Begin VB.Menu Player_Control_Menu 
         Caption         =   "&Player Control"
      End
      Begin VB.Menu Custom_Game_Menu 
         Caption         =   "Customize &Game"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Help_Content 
         Caption         =   "&Content"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu History_Arcade_Menu 
         Caption         =   "&History of Arcade"
         Enabled         =   0   'False
      End
      Begin VB.Menu Space3 
         Caption         =   "-"
      End
      Begin VB.Menu Quick_Help 
         Caption         =   "&Quick Help..."
         Enabled         =   0   'False
      End
      Begin VB.Menu Space4 
         Caption         =   "-"
      End
      Begin VB.Menu About_Nibbles 
         Caption         =   "&About Nibbles"
      End
   End
End
Attribute VB_Name = "SnakeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Key1_Up, Key1_Down, Key1_Left, Key1_Right As Integer
Public Key2_Up, Key2_Down, Key2_Left, Key2_Right As Integer
Public WaitKey1 As Boolean, WaitKey2 As Boolean

Private Sub About_Nibbles_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub Custom_Game_Menu_Click()
    Load Game_Custom
    Game_Custom.Show
End Sub

Private Sub ExitGame_Click()
    'Exit the Game upon request
    End
End Sub

Private Sub Form_Initialize()
    Me.Show
End Sub

Private Sub Form_Load()
SnakeEngine.BoxHeight = PlayScreen.Height / ArenaCol
SnakeEngine.BoxWidth = PlayScreen.Width / ArenaRow
StatusBar.Panels.Item(1).Width = 1000
StatusBar.Panels.Add.Width = 1000
StatusBar.Panels.Add.Width = 3000
StatusBar.Panels.Add.Width = 3000

'Default Variable Initiation
Key1_Up = 38
Key1_Down = 40
Key1_Left = 37
Key1_Right = 39
Key2_Up = 87
Key2_Down = 88
Key2_Left = 65
Key2_Right = 68
SnakeEngine.InitialLevel = 1
SnakeEngine.GameSpeed = 1
SnakeEngine.InitialLive = 5
SnakeEngine.ScoreBeforeLive = 3000

'--End Temp Variable Init--
End Sub

Private Sub P1_Game_Click()
SnakeEngine.PlayerNum = 1
InitGame
End Sub

Private Sub P2_Game_Click()
SnakeEngine.PlayerNum = 2
InitGame
End Sub

Private Sub Pause_Game_Click()

If Not Pause_Game.Checked Then
    MsgDraw "Paused", "Press F3 to resume", "Press F2 to start a new Game"
End If
Pause

End Sub

Private Sub Player_Control_Menu_Click()
'Activation of Player Control setup
    Load Player_Control
    Player_Control.Show
End Sub

Private Sub PlayScreen_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  '--Player 1 Keys
  Case Key1_Up                                  'Up Arror Player 1 Pressed
    If WaitKey1 Then Exit Sub
    SnakeEngine.Snake(1).DirX = 0
    If SnakeEngine.Snake(1).DirY = 1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(1).DirY = -1
    End If
    WaitKey1 = True
    
  Case Key1_Down                                'Down Arror Player 1 Pressed
    If WaitKey1 Then Exit Sub
    SnakeEngine.Snake(1).DirX = 0
    If SnakeEngine.Snake(1).DirY = -1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(1).DirY = 1
    End If
    WaitKey1 = True
    
  Case Key1_Left                                'Left Arror Player 1 Pressed
    If WaitKey1 Then Exit Sub
    SnakeEngine.Snake(1).DirY = 0
    If SnakeEngine.Snake(1).DirX = 1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(1).DirX = -1
    End If
    WaitKey1 = True
    
  Case Key1_Right                               'Right Arror Player 1 Pressed
    If WaitKey1 Then Exit Sub
    SnakeEngine.Snake(1).DirY = 0
    If SnakeEngine.Snake(1).DirX = -1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(1).DirX = 1
    End If
    WaitKey1 = True
        
 '--Player 2 Keys
 Case Key2_Up                                   'Up Arror Player 2 Pressed
    If WaitKey2 Then Exit Sub
    SnakeEngine.Snake(2).DirX = 0
    
    If SnakeEngine.Snake(2).DirY = 1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(2).DirY = -1
    End If
    WaitKey2 = True
    
  Case Key2_Down                                'Down Arror Player 2 Pressed
    If WaitKey2 Then Exit Sub
    SnakeEngine.Snake(2).DirX = 0
    If SnakeEngine.Snake(2).DirY = -1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(2).DirY = 1
    End If
    WaitKey2 = True
    
  Case Key2_Left                                'Left Arror Player 2 Pressed
    If WaitKey2 Then Exit Sub
    SnakeEngine.Snake(2).DirY = 0
    If SnakeEngine.Snake(2).DirX = 1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(2).DirX = -1
    End If
    WaitKey2 = True
    
  Case Key2_Right                               'Right Arror Player 2 Pressed
    If WaitKey2 Then Exit Sub
    SnakeEngine.Snake(2).DirY = 0
    If SnakeEngine.Snake(2).DirX = -1 Then
        Exit Sub
    Else
        SnakeEngine.Snake(2).DirX = 1
    End If
    WaitKey2 = True
End Select
End Sub

Private Sub Timer1_Timer()
'-- Check for Score, if more than ScoreBeforeLive variable, then score
'will be deducted and live would be increased--
For SnakeNo = 1 To SnakeEngine.PlayerNum
    If Snake(SnakeNo).Score >= SnakeEngine.ScoreBeforeLive Then
        Snake(SnakeNo).Score = Snake(SnakeNo).Score - SnakeEngine.ScoreBeforeLive
        Snake(SnakeNo).Lives = Snake(SnakeNo).Lives + 1
    End If
Next SnakeNo

'-- Print Score on Status Bar --
For ScorePrint = 1 To SnakeEngine.PlayerNum
    StatusBar.Panels.Item(3 + (ScorePrint - 1)).Text = SnakeEngine.Snake(ScorePrint).Score
Next ScorePrint

'--Draw Number/Points if not found--
Randomize Timer
If Not SnakeEngine.NumberOnScreen Then
    If SnakeEngine.CurNumber = 9 Then FinishLevel: Exit Sub
    Do
        NumberX = Int(Rnd(Timer) * ArenaRow) Mod ArenaRow
        NumberY = Int(Rnd(Timer) * ArenaCol) Mod ArenaCol
        If (NumberX > 0 And NumberX <= ArenaRow) And (NumberY > 0 And NumberY <= ArenaCol) Then
            If Arena(NumberX, NumberY) = Arena_None Then
                SnakeEngine.CurNumber = SnakeEngine.CurNumber + 1
                SetArena NumberX, NumberY, Arena_Number
                SnakeEngine.NumberOnScreen = True
            End If
        End If
    Loop Until SnakeEngine.NumberOnScreen
End If

'--Snake Movement--
For SnakeNum = 1 To SnakeEngine.PlayerNum
        If Not SnakeEngine.Snake(SnakeNum).Alive Then Exit For
        '-- Decide Which point is next by calculating the Position + Direction
        Newx = SnakeEngine.Snake(SnakeNum).HeadX + SnakeEngine.Snake(SnakeNum).DirX
            If Newx <= 0 Then Newx = ArenaRow
            If Newx > ArenaRow Then Newx = 1
        newy = SnakeEngine.Snake(SnakeNum).HeadY + SnakeEngine.Snake(SnakeNum).DirY
            If newy <= 0 Then newy = ArenaCol
            If newy > ArenaCol Then newy = 1
        
        '-- See whats on the point ahead (eg. walls or points)
        If Arena(Newx, newy) = Arena_Wall Then SnakeDead (SnakeNum): Exit Sub
        If Arena(Newx, newy) = Arena_Snake1 Then SnakeDead (SnakeNum): Exit Sub
        If Arena(Newx, newy) = Arena_Snake2 Then SnakeDead (SnakeNum): Exit Sub
        If Arena(Newx, newy) = Arena_Number Then
            New_Snake_Length = SnakeEngine.Snake(SnakeNum).Length + 3 + SnakeEngine.CurNumber
            SnakeEngine.Snake(SnakeNum).Score = SnakeEngine.Snake(SnakeNum).Score + SnakeEngine.CurNumber * 20
            If New_Snake_Length > SnakeEngine.MaxSnakeLength Then New_Snake_Length = SnakeEngine.MaxSnakeLength
            SnakeEngine.Snake(SnakeNum).Length = New_Snake_Length
            SnakeEngine.NumberOnScreen = False
        End If
        WaitKey1 = False: WaitKey2 = False
        '-- Colour the Snake
        If SnakeNum = 1 Then SnakeColor = Arena_Snake1
        If SnakeNum = 2 Then SnakeColor = Arena_Snake2
        SnakeAddBody SnakeEngine.Snake(SnakeNum), Newx, newy, SnakeColor
Next SnakeNum
End Sub

Private Sub WireFrame_Click()
If SnakeFrm.WireFrame.Checked Then
    SnakeFrm.WireFrame.Checked = False
Else
    SnakeFrm.WireFrame.Checked = True
End If
End Sub
