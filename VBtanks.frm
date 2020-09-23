VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tanks"
   ClientHeight    =   4425
   ClientLeft      =   1665
   ClientTop       =   1500
   ClientWidth     =   7365
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "VBTANKS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   Tag             =   "0"
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar BluePower 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      Max             =   1000
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.HScrollBar BlueAngle 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      Max             =   180
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.HScrollBar RedAngle 
      Height          =   255
      Left            =   120
      Max             =   0
      Min             =   180
      TabIndex        =   2
      Top             =   480
      Value           =   180
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Fire 
      Caption         =   "FIRE"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.HScrollBar RedPower 
      Height          =   255
      Left            =   120
      Max             =   1000
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Angle"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Power"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Power"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Angle"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image tank2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6120
      Picture         =   "VBTANKS.frx":030A
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image tank1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   480
      Picture         =   "VBTANKS.frx":0BD4
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Angle_Change()

' If it's red's turn then red angle changes
If Form1.Tag = "red" Then Form1.RedAngle = Angle

' If it's blue's turn then blue angle changes
If Form1.Tag = "blue" Then Form1.BlueAngle = Angle

End Sub

Private Sub Angle_Scroll()

' If it's red's turn then red angle changes
If Form1.Tag = "red" Then Form1.RedAngle = Angle

' If it's blue's turn then blue angle changes
If Form1.Tag = "blue" Then Form1.BlueAngle = Angle

End Sub


Private Sub Command1_Click()

End Sub

Private Sub BlueAngle_Change()
Label4.Caption = "Angle: " & BlueAngle
draw_gun
End Sub

Private Sub BlueAngle_Scroll()
Label4.Caption = "Angle: " & BlueAngle
draw_gun
End Sub


Private Sub BluePower_Change()
Label3.Caption = "Power: " & BluePower
End Sub

Private Sub BluePower_Scroll()
Label3.Caption = "Power: " & BluePower
End Sub


Private Sub Fire_Click()

' If it's red's turn then red fires
If Form1.Tag = "red" Then
red_shot
BlueAngle.Enabled = True
BluePower.Enabled = True
RedAngle.Enabled = False
RedPower.Enabled = False
Form1.Tag = "blue"
Exit Sub
End If

' If it's blue's turn then blue fires
If Form1.Tag = "blue" Then
blue_shot
BlueAngle.Enabled = False
BluePower.Enabled = False
RedAngle.Enabled = True
RedPower.Enabled = True
Form1.Tag = "red"
Exit Sub
End If

End Sub


Private Sub Picture1_Click()

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub Power_Change()

' If it's red's turn then red angle changes
If Form1.Tag = "red" Then Form1.RedPower = Power

' If it's blue's turn then blue angle changes
If Form1.Tag = "blue" Then Form1.BluePower = Power

End Sub

Private Sub Power_Scroll()

' If it's red's turn then red angle changes
If Form1.Tag = "red" Then Form1.RedPower = Power

' If it's blue's turn then blue angle changes
If Form1.Tag = "blue" Then Form1.BluePower = Power

End Sub


Private Sub Form_Resize()
landscape2 Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
' MsgBox "This game was made by Haig Demerdjian. Thanks for playing!"
End Sub

Private Sub RedAngle_Change()
Label1.Caption = "Angle: " & RedAngle
draw_gun
End Sub


Private Sub RedAngle_Scroll()
Label1.Caption = "Angle: " & RedAngle
draw_gun
End Sub


Private Sub RedPower_Change()
Label2.Caption = "Power: " & RedPower
End Sub


Private Sub RedPower_Scroll()
Label2.Caption = "Power: " & RedPower
End Sub


