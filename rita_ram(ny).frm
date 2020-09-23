VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rossing Paint"
   ClientHeight    =   10650
   ClientLeft      =   975
   ClientTop       =   45
   ClientWidth     =   15240
   Icon            =   "rita_ram(ny).frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "rita_ram(ny).frx":0442
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "New"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "Print"
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu window 
      Caption         =   "&Window"
      Begin VB.Menu tools 
         Caption         =   "Tools"
      End
      Begin VB.Menu paint 
         Caption         =   "Paint"
      End
      Begin VB.Menu tool 
         Caption         =   "Tool Properties"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cordtinate_Click()
Form4.Show
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()
  
 '
  Form2.Show
  Form1.Show
  Form3.Show
  Form4.Show

End Sub

Private Sub MDIForm_Resize()
'On Error GoTo zize
'Me.Width = 15480
'Me.Height = 12345
'zize:
'Me.WindowState = 0

End Sub

Private Sub new_Click()
Form5.Show
End Sub

Private Sub open_Click()
On Error GoTo 10
    Form1.CommonDialog1.ShowOpen
    Form3.Pic1.Picture = LoadPicture(Form1.CommonDialog1.FileName)
10
End Sub

Private Sub paint_Click()
Form3.Show
End Sub

Private Sub print_Click()
'On Error GoTo 9129

SavePicture Form3.Pic1.Image, ("C:\bas.bmp")
Form3.Pic1.Picture = LoadPicture("C:\bas.bmp")
Form1.CommonDialog1.ShowPrinter
Printer.PaintPicture Form3.Pic1.Picture, 0, 0

Printer.EndDoc

'9129
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub save_Click()
   Form1.CommonDialog3.InitDir = App.Path
    Form1.CommonDialog3.Filter = "Bitmap Image (*.bmp)|*.bmp|jpeg Image (*.jpg)|*.jpg|Bitmap Image (*.gif)|*.gif|"
    Form1.CommonDialog3.DialogTitle = "Save Rossing Paint"
    Form1.CommonDialog3.ShowSave
If Form1.CommonDialog3.FileName <> "" Then
    SavePicture Form3.Pic1.Image, Form1.CommonDialog3.FileName
End If
End Sub

Private Sub tool_Click()
Form1.Show
End Sub

Private Sub tools_Click()
Form2.Show
End Sub
Sub sice()

End Sub
