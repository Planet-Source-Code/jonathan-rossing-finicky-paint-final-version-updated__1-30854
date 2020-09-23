VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9555
   ClientLeft      =   3600
   ClientTop       =   3780
   ClientWidth     =   2805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Rita4ny).frx":0000
   ScaleHeight     =   9555
   ScaleMode       =   0  'User
   ScaleWidth      =   1.958
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   10
      Left            =   270
      Picture         =   "Rita4ny).frx":8A13
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "#D rör"
      Top             =   1140
      Width           =   105
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   1
      Left            =   180
      Picture         =   "Rita4ny).frx":8E55
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "#D rör"
      Top             =   825
      Width           =   180
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   165
      Picture         =   "Rita4ny).frx":8F82
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Rita"
      Top             =   1275
      Value           =   -1  'True
      Width           =   120
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   3
      Left            =   195
      Picture         =   "Rita4ny).frx":909E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Bakgrunds färg"
      Top             =   975
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   4
      Left            =   210
      Picture         =   "Rita4ny).frx":91AA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Rektangel"
      Top             =   345
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   5
      Left            =   210
      Picture         =   "Rita4ny).frx":9339
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Hämta färg"
      Top             =   675
      Width           =   135
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   6
      Left            =   285
      Picture         =   "Rita4ny).frx":9475
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Linje fjäder"
      Top             =   1305
      Width           =   120
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   7
      Left            =   165
      Picture         =   "Rita4ny).frx":95B1
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Linje"
      Top             =   1470
      Width           =   105
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   8
      Left            =   285
      Picture         =   "Rita4ny).frx":99F3
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Stjärn mönster"
      Top             =   1485
      Width           =   105
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   9
      Left            =   135
      Picture         =   "Rita4ny).frx":9B2F
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Bubblor"
      Top             =   1080
      Width           =   120
      Visible         =   0   'False
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Index           =   0
      Left            =   225
      Picture         =   "Rita4ny).frx":B9E7
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cirkel"
      Top             =   510
      Width           =   150
      Visible         =   0   'False
   End
   Begin VB.Image Image11 
      Height          =   1170
      Left            =   1395
      Picture         =   "Rita4ny).frx":BB14
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Image Image10 
      Height          =   960
      Left            =   1560
      Picture         =   "Rita4ny).frx":E14C
      Top             =   6900
      Width           =   1020
   End
   Begin VB.Image Image9 
      Height          =   900
      Left            =   1245
      Picture         =   "Rita4ny).frx":103BF
      Top             =   5610
      Width           =   975
   End
   Begin VB.Image Image8 
      Height          =   930
      Left            =   150
      Picture         =   "Rita4ny).frx":10D7B
      Top             =   4605
      Width           =   1125
   End
   Begin VB.Image Image7 
      Height          =   1290
      Left            =   135
      Picture         =   "Rita4ny).frx":12E65
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Image Image6 
      Height          =   930
      Left            =   1560
      Picture         =   "Rita4ny).frx":151DC
      Top             =   3090
      Width           =   990
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   1485
      Picture         =   "Rita4ny).frx":17122
      Top             =   4275
      Width           =   1035
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   165
      Picture         =   "Rita4ny).frx":1946A
      Top             =   3360
      Width           =   990
   End
   Begin VB.Image Image3 
      Height          =   1185
      Left            =   1260
      Picture         =   "Rita4ny).frx":1B857
      Top             =   90
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   240
      Picture         =   "Rita4ny).frx":1DD95
      Top             =   150
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   75
      Picture         =   "Rita4ny).frx":1FAAF
      Top             =   5730
      Width           =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move 0, 1950
End Sub

Private Sub Image1_Click()
Option1(9) = True
End Sub

Private Sub Image10_Click()
Option1(8) = True
End Sub

Private Sub Image11_Click()
Option1(7) = True
End Sub

Private Sub Image2_Click()
Option1(3) = True
End Sub

Private Sub Image3_Click()
Option1(2) = True
End Sub

Private Sub Image4_Click()
Option1(4) = True
End Sub

Private Sub Image5_Click()
Option1(0) = True
End Sub

Private Sub Image6_Click()
Option1(5) = True
End Sub

Private Sub Image7_Click()
Option1(1) = True
End Sub

Private Sub Image8_Click()
Option1(10) = True
End Sub

Private Sub Image9_Click()
Option1(6) = True
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(1).Value = True Then
    Form1.Frame1(0).Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame1(0).Visible = False
End If
If Option1(2).Value = True Then
    Form1.Frame2.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame2.Visible = False
End If
If Option1(3).Value = True Then
    Form1.Frame3.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
    Form3.Pic1.MousePointer = 1
Else
    Form1.Frame3.Visible = False
    Form3.Pic1.MousePointer = 99
End If
If Option1(4).Value = True Then
    Form1.Frame4.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame4.Visible = False
End If
If Option1(5).Value = True Then
    Form1.Frame5.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame5.Visible = False
End If
If Option1(6).Value = True Then
    Form1.Frame6.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame6.Visible = False
End If
If Option1(7).Value = True Then
    Form1.Frame7.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame7.Visible = False
End If
If Option1(8).Value = True Then
    Form1.Frame8.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame8.Visible = False
End If
If Option1(9).Value = True Then
    Form1.Frame9.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
    Form1.BackColor = Farg
Else
    Form1.Frame9.Visible = False
End If
If Option1(0).Value = True Then
    Form1.Frame10.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame10.Visible = False
End If
If Option1(10).Value = True Then
    Form1.Frame12.Visible = True
    Form3.Pic1.FillStyle = 0
    'Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame12.Visible = False
End If
End Sub

Private Sub rgbs_Change()

End Sub
