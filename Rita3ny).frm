VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Tool Box"
   ClientHeight    =   1515
   ClientLeft      =   12135
   ClientTop       =   2070
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Rita3ny).frx":0000
   ScaleHeight     =   1515
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Flood fill"
      Height          =   1005
      Left            =   225
      TabIndex        =   39
      Top             =   285
      Width           =   2370
      Visible         =   0   'False
      Begin VB.HScrollBar redScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   0
         Max             =   255
         TabIndex        =   43
         Top             =   480
         Width           =   735
      End
      Begin VB.HScrollBar greenScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   795
         Max             =   255
         TabIndex        =   42
         Top             =   480
         Width           =   735
      End
      Begin VB.HScrollBar blueScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   1560
         Max             =   255
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Color"
         Height          =   240
         Left            =   480
         TabIndex        =   40
         Top             =   645
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   795
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   45
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   1080
         Top             =   645
         Width           =   270
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   7005
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   36
      Top             =   4320
      Width           =   15
      Visible         =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eraser"
      Height          =   675
      Left            =   330
      TabIndex        =   35
      Top             =   450
      Width           =   2130
      Visible         =   0   'False
      Begin VB.Label N 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "No controls"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3D pipe"
      Height          =   1095
      Index           =   0
      Left            =   600
      TabIndex        =   30
      Top             =   255
      Width           =   1695
      Visible         =   0   'False
      Begin VB.CommandButton Command4 
         Caption         =   "Color"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fill color"
         Height          =   315
         Left            =   330
         TabIndex        =   32
         Top             =   675
         Width           =   1035
      End
      Begin VB.HScrollBar HS2 
         Height          =   150
         LargeChange     =   10
         Left            =   840
         Max             =   100
         Min             =   7
         TabIndex        =   31
         Top             =   480
         Value           =   7
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penn"
      Height          =   720
      Left            =   600
      TabIndex        =   26
      Top             =   240
      Width           =   1695
      Visible         =   0   'False
      Begin VB.CommandButton Command5 
         Caption         =   "Color"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.HScrollBar HS1 
         Height          =   165
         Left            =   840
         Max             =   400
         Min             =   1
         TabIndex        =   27
         Top             =   480
         Value           =   1
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rectangle"
      Height          =   945
      Left            =   615
      TabIndex        =   22
      Top             =   300
      Width           =   1695
      Visible         =   0   'False
      Begin VB.HScrollBar rgbs2 
         Height          =   150
         Left            =   855
         Max             =   256
         TabIndex        =   50
         Top             =   450
         Width           =   780
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Color"
         Height          =   210
         Left            =   270
         TabIndex        =   24
         Top             =   645
         Width           =   1095
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   435
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   318
         _Version        =   327682
         LargeChange     =   2
         Min             =   1
         Max             =   50
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "color"
         Height          =   165
         Left            =   930
         TabIndex        =   51
         Top             =   255
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   165
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get color"
      Height          =   900
      Left            =   720
      TabIndex        =   20
      Top             =   240
      Width           =   1575
      Visible         =   0   'False
      Begin VB.CommandButton Command8 
         Caption         =   "Make it so"
         Height          =   495
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   735
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   1140
         Top             =   270
         Width           =   315
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   1140
         Top             =   570
         Width           =   315
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sun feather"
      ClipControls    =   0   'False
      DragIcon        =   "Rita3ny).frx":2EBD
      DragMode        =   1  'Automatic
      Height          =   735
      Left            =   720
      TabIndex        =   17
      Top             =   240
      Width           =   1515
      Visible         =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "Color"
         Height          =   255
         Left            =   765
         TabIndex        =   38
         Top             =   315
         Width           =   645
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   105
         Max             =   10
         Min             =   1
         TabIndex        =   18
         Top             =   405
         Value           =   1
         Width           =   585
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stamp"
      Height          =   930
      Left            =   600
      TabIndex        =   16
      Top             =   240
      Width           =   1605
      Visible         =   0   'False
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "No controls"
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stars"
      Height          =   960
      Left            =   840
      TabIndex        =   12
      Top             =   240
      Width           =   1395
      Visible         =   0   'False
      Begin VB.CommandButton Command11 
         Caption         =   "Color"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   210
         Width           =   975
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   180
         Left            =   195
         Max             =   300
         Min             =   5
         TabIndex        =   13
         Top             =   720
         Value           =   5
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bubbel Spray"
      Height          =   1140
      Left            =   300
      TabIndex        =   3
      Top             =   210
      Width           =   2175
      Visible         =   0   'False
      Begin VB.CommandButton Command12 
         Caption         =   "Fill color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   975
         TabIndex        =   8
         Top             =   615
         Width           =   750
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Color"
         Height          =   210
         Left            =   975
         TabIndex        =   7
         Top             =   855
         Width           =   750
      End
      Begin ComctlLib.Slider Slider4 
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   870
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   344
         _Version        =   327682
         Min             =   1
         Max             =   30
         SelStart        =   1
         Value           =   1
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   390
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   344
         _Version        =   327682
         LargeChange     =   2
         Min             =   5
         Max             =   30
         SelStart        =   5
         Value           =   5
      End
      Begin ComctlLib.Slider S1 
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   420
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         _Version        =   327682
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Distribution"
         Height          =   225
         Left            =   1080
         TabIndex        =   10
         Top             =   165
         Width           =   900
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Amount "
         Height          =   180
         Left            =   165
         TabIndex        =   9
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ovel"
      Height          =   1095
      Left            =   405
      TabIndex        =   0
      Top             =   195
      Width           =   1965
      Visible         =   0   'False
      Begin VB.HScrollBar rgbs 
         Height          =   195
         Left            =   945
         Max             =   256
         TabIndex        =   47
         Top             =   480
         Width           =   885
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Color"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin ComctlLib.Slider s2 
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   318
         _Version        =   327682
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Color"
         Height          =   225
         Left            =   945
         TabIndex        =   49
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label9 
         Caption         =   "Border"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   525
      Left            =   195
      TabIndex        =   37
      Top             =   240
      Width           =   525
      Visible         =   0   'False
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   15
         Top             =   15
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog4 
         Left            =   0
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   15
         Top             =   15
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   15
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Open..."
         Filter          =   "Image Files (*.bmp)|*.bmp|jpeg bilder(*.jpg)|*.jpg|gif Files (*.gif)|*.gif|"
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   870
         Width           =   375
      End
      Begin VB.Shape Shape4 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   1860
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   2310
         Width           =   375
      End
      Begin VB.Shape Shape6 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2940
         Top             =   1380
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Färgen, fargen, Farg, farg2, linjen As Double
Dim tryck As Boolean

 Sub blueScroll_Change()
    Label5.Caption = blueScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)

End Sub

Public Sub Form_Click()


End Sub

Public Sub Form_GotFocus()


End Sub

Private Sub Form_Load()
Me.Move 9300, 1900

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Public Sub Frame11_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Public Sub greenScroll_Change()
    Label6.Caption = greenScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)

End Sub

Public Sub redscroll_Change()
    Label4.Caption = redScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)
End Sub

Public Sub pCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    colpressed = True
    Shape8.FillColor = pCol.Point(X, Y)
  
End Sub

Public Sub HS1_Change()
    Label2.Caption = HS1.Value
End Sub

Public Sub HS2_Change()
    Label1.Caption = HS2.Value
End Sub

Public Sub jg_Click()

End Sub

Public Sub HScroll1_Change()
    Label14.Caption = HScroll1.Value
    Form3.Pic1.DrawWidth = HScroll1.Value
End Sub
Public Sub Command1_Click()
   CommonDialog1.ShowColor
End Sub

Private Sub Command10_Click()
On Error GoTo 5544
    CommonDialog1.ShowColor
   
5544
End Sub

Private Sub Command11_Click()
    CommonDialog3.ShowColor
    
End Sub

Private Sub Command12_Click()
 CommonDialog2.ShowColor
    
End Sub

Private Sub Command13_Click()
On Error GoTo 6655
    CommonDialog1.ShowColor
    
6655
End Sub

Private Sub Command14_Click()
On Error GoTo 5566
    CommonDialog1.ShowColor
   
5566
End Sub

Private Sub Command2_Click()
On Error GoTo 5665
CommonDialog1.ShowColor
5665
End Sub

Private Sub Command3_Click()
    CommonDialog2.ShowColor
    
End Sub

Private Sub Command4_Click()
On Error GoTo 321
    CommonDialog1.ShowColor
    
321
End Sub

Private Sub Command5_Click()
    CommonDialog3.ShowColor
    
End Sub



Private Sub Command7_Click()
    CommonDialog3.ShowColor
    
End Sub


Public Sub Command8_Click()
    farg2 = Shape5.FillColor
          fargen = Form1.Shape5.FillColor
     farg2 = Form1.Shape5.FillColor
     Färgen = Form1.Shape5.FillColor
    Farg = Form1.Shape5.FillColor

End Sub

