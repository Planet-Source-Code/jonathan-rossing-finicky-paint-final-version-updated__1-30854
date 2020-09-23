VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Cordinates"
   ClientHeight    =   1515
   ClientLeft      =   7485
   ClientTop       =   780
   ClientWidth     =   2985
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Rita(ny).frx":0000
   ScaleHeight     =   1515
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   3
      Top             =   795
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   795
      Width           =   960
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X-axel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   300
      TabIndex        =   1
      Top             =   435
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y-axel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Top             =   435
      Width           =   1020
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move 12400, 1900
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
