VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Game_Custom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Option"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame GameSetFrame 
      Caption         =   "Game &Setting"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1000
         Increment       =   1000
         Max             =   500000
         Min             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.TextBox PointLiveLabel 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3015
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   9
      End
      Begin VB.Label LevelLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2670
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label SpeedLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2675
         TabIndex        =   10
         Top             =   1900
         Width           =   375
      End
      Begin VB.Label LiveLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2675
         TabIndex        =   6
         Top             =   340
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Initial &Level"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Game &Speed"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Points needed for &Bonus Live"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Number of &Lives"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "Game_Custom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SnakeEngine.InitialLive = Slider1.Value
SnakeEngine.InitialLevel = Slider3.Value
SnakeEngine.GameSpeed = (10 - Slider2.Value) * 10 + 1
SnakeEngine.ScoreBeforeLive = UpDown1.Value
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Slider1.Value = SnakeEngine.InitialLive
UpDown1.Value = SnakeEngine.ScoreBeforeLive
Slider2.Value = Int((100 - SnakeEngine.GameSpeed) / 10) + 1
Slider3.Value = SnakeEngine.InitialLevel
End Sub

Private Sub Slider1_Change()
LiveLabel.Caption = Slider1.Value
End Sub

Private Sub Slider2_Change()
SpeedLabel.Caption = Slider2.Value
End Sub


Private Sub Slider3_Change()
LevelLabel.Caption = Slider3.Value
End Sub


Private Sub UpDown1_Change()
PointLiveLabel.Text = UpDown1.Value
End Sub
