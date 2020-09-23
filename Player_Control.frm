VERSION 5.00
Begin VB.Form Player_Control 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize Control ..."
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel_Btn 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OK_Btn 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Player2_Frame 
      Caption         =   "Player 2 Control"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
      Begin VB.ComboBox P2_Up 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox P2_Down 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox P2_Left 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox P2_Right 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Up"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Down"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Left"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Right"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Player1_Frame 
      Caption         =   "Player 1 Control"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox P1_Right 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox P1_Left 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox P1_Down 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox P1_Up 
         Height          =   315
         ItemData        =   "Player_Control.frx":0000
         Left            =   720
         List            =   "Player_Control.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Right"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Left"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Down"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Up"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   210
      End
   End
End
Attribute VB_Name = "Player_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyCodeString(1 To 95) As String
Dim KeyCodeNum(1 To 95) As Integer
Dim ComboSetIndex(1 To 8) As Integer

Private Sub Cancel_Btn_Click()
Unload Me
End Sub

Private Sub Form_Load()
'-- Initialization Key Completed --
KeyCodeNum(1) = 8
KeyCodeString(1) = "Back Space"
KeyCodeNum(2) = 9
KeyCodeString(2) = "Tab"
KeyCodeNum(3) = 13
KeyCodeString(3) = "Enter"
KeyCodeNum(4) = 16
KeyCodeString(4) = "Shift"
KeyCodeNum(5) = 17
KeyCodeString(5) = "Ctrl"
KeyCodeNum(6) = 18
KeyCodeString(6) = "Alt"
KeyCodeNum(7) = 19
KeyCodeString(7) = "Pause Break"
KeyCodeNum(8) = 20
KeyCodeString(8) = "Caps Lock"
KeyCodeNum(9) = 18
KeyCodeString(9) = "Alt"
KeyCodeNum(10) = 27
KeyCodeString(10) = "Esc"
KeyCodeNum(11) = 38
KeyCodeString(11) = "Up Arrow"
KeyCodeNum(12) = 40
KeyCodeString(12) = "Down Arrow"
KeyCodeNum(13) = 37
KeyCodeString(13) = "Left Arrow"
KeyCodeNum(14) = 39
KeyCodeString(14) = "Right Arrow"
KeyCodeNum(15) = 45
KeyCodeString(15) = "Ins"
KeyCodeNum(16) = 46
KeyCodeString(16) = "Del"
KeyCodeNum(17) = 36
KeyCodeString(17) = "Home"
KeyCodeNum(18) = 35
KeyCodeString(18) = "End"
KeyCodeNum(19) = 33
KeyCodeString(19) = "PgUp"
KeyCodeNum(20) = 34
KeyCodeString(20) = "PgDown"
For NumPad = 0 To 9
    KeyCodeNum(21 + NumPad) = 96 + NumPad
    KeyCodeString(21 + NumPad) = "Numpad " + LTrim$(Str$(NumPad))
Next NumPad
KeyCodeNum(31) = 110
KeyCodeString(31) = "Numpad ."
KeyCodeNum(32) = 106
KeyCodeString(32) = "Numpad *"
KeyCodeNum(33) = 111
KeyCodeString(33) = "Numpad /"
KeyCodeNum(34) = 109
KeyCodeString(34) = "Numpad -"
KeyCodeNum(35) = 107
KeyCodeString(35) = "Numpad +"
For Number = 0 To 9
    KeyCodeNum(36 + NumPad) = 48 + NumPad
    KeyCodeString(36 + NumPad) = LTrim$(Str$(NumPad))
Next Number
KeyCodeNum(46) = 192
KeyCodeString(46) = "`"
KeyCodeNum(47) = 189
KeyCodeString(47) = "-"
KeyCodeNum(48) = 187
KeyCodeString(48) = "="
KeyCodeNum(49) = 220
KeyCodeString(49) = "\"
For Character = 65 To 90
    KeyCodeNum(Character - 65 + 50) = Character
    KeyCodeString(Character - 65 + 50) = Chr$(Character)
Next Character
KeyCodeNum(76) = 219
KeyCodeString(76) = "["
KeyCodeNum(77) = 221
KeyCodeString(77) = "]"
KeyCodeNum(78) = 186
KeyCodeString(78) = ";"
KeyCodeNum(79) = 222
KeyCodeString(79) = "'"
KeyCodeNum(80) = 188
KeyCodeString(80) = ","
KeyCodeNum(81) = 190
KeyCodeString(81) = "."
KeyCodeNum(82) = 191
KeyCodeString(82) = "/"
KeyCodeNum(83) = 32
KeyCodeString(83) = "Space Bar"
For Fes = 1 To 12
    KeyCodeNum(83 + Fes) = 111 + Fes
    KeyCodeString(83 + Fes) = "F" + LTrim$(Str$(Fes))
Next Fes
'-- Initialization Key Completed --

'-- Set Key at ComboBox --
For KeyCodeNo = 1 To 95
    P1_Up.AddItem KeyCodeString(KeyCodeNo)
    P1_Down.AddItem KeyCodeString(KeyCodeNo)
    P1_Left.AddItem KeyCodeString(KeyCodeNo)
    P1_Right.AddItem KeyCodeString(KeyCodeNo)
    
    P2_Up.AddItem KeyCodeString(KeyCodeNo)
    P2_Down.AddItem KeyCodeString(KeyCodeNo)
    P2_Left.AddItem KeyCodeString(KeyCodeNo)
    P2_Right.AddItem KeyCodeString(KeyCodeNo)
Next KeyCodeNo
'-- Set Key End --

'-- Position Current Key at the ComboBox --
For KeyCodeNo = 1 To 95
    If SnakeFrm.Key1_Up = KeyCodeNum(KeyCodeNo) Then P1_Up.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key1_Down = KeyCodeNum(KeyCodeNo) Then P1_Down.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key1_Left = KeyCodeNum(KeyCodeNo) Then P1_Left.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key1_Right = KeyCodeNum(KeyCodeNo) Then P1_Right.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key2_Up = KeyCodeNum(KeyCodeNo) Then P2_Up.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key2_Down = KeyCodeNum(KeyCodeNo) Then P2_Down.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key2_Left = KeyCodeNum(KeyCodeNo) Then P2_Left.ListIndex = KeyCodeNo - 1
    If SnakeFrm.Key2_Right = KeyCodeNum(KeyCodeNo) Then P2_Right.ListIndex = KeyCodeNo - 1
Next KeyCodeNo
'-- End ComboBox Positioning--
End Sub

Private Sub OK_Btn_Click()
ComboSetIndex(1) = P1_Up.ListIndex
ComboSetIndex(2) = P1_Down.ListIndex
ComboSetIndex(3) = P1_Left.ListIndex
ComboSetIndex(4) = P1_Right.ListIndex
ComboSetIndex(5) = P2_Up.ListIndex
ComboSetIndex(6) = P2_Down.ListIndex
ComboSetIndex(7) = P2_Left.ListIndex
ComboSetIndex(8) = P2_Right.ListIndex
For i = 1 To 8
    For j = 1 To 8
        If i = j Then Exit For
        If ComboSetIndex(i) = ComboSetIndex(j) Then
            KeyMsg$ = "One or more keys has been specified for multiple function." + Chr$(13) + "Please modify the key configuration to prevent conflicts."
            MsgBox KeyMsg$, vbCritical + vbOKOnly, "Key Conflict"
            Exit Sub
        End If
    Next j
Next i
SnakeFrm.Key1_Up = KeyCodeNum(ComboSetIndex(1) + 1)
SnakeFrm.Key1_Down = KeyCodeNum(ComboSetIndex(2) + 1)
SnakeFrm.Key1_Left = KeyCodeNum(ComboSetIndex(3) + 1)
SnakeFrm.Key1_Right = KeyCodeNum(ComboSetIndex(4) + 1)
SnakeFrm.Key2_Up = KeyCodeNum(ComboSetIndex(5) + 1)
SnakeFrm.Key2_Down = KeyCodeNum(ComboSetIndex(6) + 1)
SnakeFrm.Key2_Left = KeyCodeNum(ComboSetIndex(7) + 1)
SnakeFrm.Key2_Right = KeyCodeNum(ComboSetIndex(8) + 1)
Unload Me
End Sub
