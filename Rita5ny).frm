VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "New"
   ClientHeight    =   4410
   ClientLeft      =   2025
   ClientTop       =   5220
   ClientWidth     =   3435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   Picture         =   "Rita5ny).frx":0000
   ScaleHeight     =   4410
   ScaleWidth      =   3435
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "3600"
      Top             =   1080
      Width           =   765
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "3600"
      Top             =   1680
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Rita5ny).frx":7C4C
      Left            =   1800
      List            =   "Rita5ny).frx":7C62
      TabIndex        =   1
      Text            =   "Color:"
      Top             =   2520
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   7
      Top             =   1710
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   855
      TabIndex        =   6
      Top             =   1125
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   630
      TabIndex        =   5
      Top             =   2610
      Width           =   1155
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Combo1.Text = "color:" Then
      MsgBox "Color?"
End If
If Combo1.Text = "Black" Then
      Form3.Pic1.BackColor = vbBlack
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Blue" Then
      Form3.Pic1.BackColor = vbBlue
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Red" Then
      Form3.Pic1.BackColor = vbRed
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Green" Then
      Form3.Pic1.BackColor = vbGreen
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "White" Then
      Form3.Pic1.BackColor = &HFFFFFF
      Form3.Pic1.Picture = LoadPicture
      End If
   
  
   Form3.Pic1.Width = Text1
   Form3.Pic1.Height = Text2
     Form3.Width = Form3.Pic1.Width
     Form3.Height = Form3.Pic1.Height
   
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_Change()
If Text2.Text > 9825 Then
    Text2.Text = 3600
    MsgBox " Max 9825"
End If
End Sub

