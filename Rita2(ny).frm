VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paint"
   ClientHeight    =   4710
   ClientLeft      =   5955
   ClientTop       =   1890
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   -15
      MouseIcon       =   "Rita2(ny).frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   -15
      Width           =   4455
      Begin VB.Image stamp 
         Height          =   1035
         Left            =   -720
         Picture         =   "Rita2(ny).frx":030A
         Top             =   -840
         Width           =   990
         Visible         =   0   'False
      End
      Begin VB.Shape Shape3 
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Left            =   120
         Shape           =   2  'Oval
         Top             =   840
         Width           =   375
         Visible         =   0   'False
      End
      Begin VB.Shape Shape4 
         DrawMode        =   6  'Mask Pen Not
         Height          =   255
         Left            =   210
         Top             =   2280
         Width           =   375
         Visible         =   0   'False
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   44
         X2              =   42
         Y1              =   220
         Y2              =   286
      End
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Top             =   810
      Width           =   375
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3870
      Top             =   1740
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3870
      Top             =   1290
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Top             =   300
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************************************************************

'                  Please give me some credit if you use my code
'**********************************J Rossing***************************************

'*********************************************************************************
Dim Index As Integer
Dim colpressed As Boolean
Dim kludda As Boolean
Dim klick As Boolean
Dim Färgen As Integer
Dim fargen As Integer
Dim Farg As Integer
Dim farg2 As Integer
Dim linjen As Integer
Dim qw As Integer
Dim radie As Integer
Dim radie2 As Integer
Dim mittfarg As Integer
Dim upp As Integer
Dim lod As Single
Dim våg As Single
Dim lod1 As Single
Dim våg1 As Single
Dim XX As Double, YY As Double
Dim XX2 As Double, YY2 As Double
Dim CurrentChoice

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then kludda = True
våg = X
    lod = Y
fargen = Form1.CommonDialog3.Color
farg2 = Form1.CommonDialog2.Color
Farg = Form1.CommonDialog1.Color

    Färgen = fargen
 Farg = Form1.CommonDialog1.Color
 qw2 = Form1.rgbs.Value
On Error GoTo 12








'***Paint bucket***'
If Form2.Option1(3).Value = True Then
    Pic1.DrawWidth = 5
    Pic1.FillColor = vbWhite


End If
'***Get color***'
If Form2.Option1(5).Value = True Then Form1.Shape5.FillColor = Pic1.Point(X, Y)
   
'***Penn***'
If kludda = True And Form2.Option1(2).Value = True Then
    Pic1.DrawWidth = HS1
     
End If

If kludda = True And Form2.Option1(10).Value = True Then
    
     
End If
    
'***Rektangle***'
If kludda = True And Form2.Option1(4).Value = True Then
     XX = X: YY = Y
     XX2 = X: YY2 = Y
     Shape4.Visible = True
     Shape4.Left = X: Shape2.Top = Y
     Shape4.Width = 0: Shape4.Height = 0
     qw = Form1.rgbs2.Value
End If

'***oval***'
 If kludda = True And Form2.Option1(0).Value = True Then
    XX = X: YY = Y
    XX2 = X: YY2 = Y
    Shape3.Visible = True
    Shape3.Left = X: Shape2.Top = Y
                    Shape3.Width = 0: Shape3.Height = 0
                    qw = Form1.rgbs.Value
End If
        
'***Line***'
If kludda = True And Form2.Option1(7).Value = True Then
 '   Line2.X1 = X: Line2.X2 = X
  '      Line2.Y1 = Y: Line2.Y2 = Y
  '          Line2.Visible = True

Pic1.PaintPicture stamp.Picture, X - 30, Y - 30

End If
'***Feather***'
If kludda And Form2.Option1(6).Value = True Then
    XX = X: YY = Y
End If

If kludda = True And Form2.Option1(10).Value = True Then
Pic1.FillColor = Form1.Shape1.FillColor
ExtFloodFill Pic1.hdc, X, Y, Pic1.Point(X, Y), 1
End If
12
End Sub

Public Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Text5 = X
    Text6 = Y
    Pic1.FillColor = farg2
    Form4.Label1 = X
    Form4.Label2 = Y

If Form2.Option1(3).Value = True And kludda = True Then
    Pic1.FillColor = vbWhite
    Pic1.Circle (X, Y), 10, vbWhite
        
End If


'***get color***'
If Form2.Option1(5).Value = True Then
    Form1.Shape2.FillColor = Pic1.Point(X, Y)
      fargen = Form1.Shape2.FillColor
     farg2 = Form1.Shape2.FillColor
     Färgen = Form1.Shape2.FillColor
    Farg = Form1.Shape2.FillColor

    
End If

'***Penn***'
If kludda And Form2.Option1(2).Value = True Then
   Pic1.DrawWidth = Form1.HS1
    Pic1.FillColor = fargen
End If



'***pipe***'
If kludda And Form2.Option1(1).Value = True Then
Pic1.Circle (X, Y), Form1.HS2, Farg
    
    Pic1.DrawWidth = 1
        Form1.HS1.Value = 1
            Färgen = farg2
            Pic1.DrawWidth = 2
            Pic1.Line (X, Y - (Form1.HS2))-(våg, lod - (Form1.HS2)), Farg
            Pic1.Line (X, Y + (Form1.HS2))-(våg, lod + (Form1.HS2)), Farg
            Pic1.Line (X - (Form1.HS2), Y)-(våg - (Form1.HS2), lod), Farg
            Pic1.Line (X + (Form1.HS2), Y)-(våg + (Form1.HS2), lod), Farg
            Pic1.DrawWidth = 1
            våg = X
            lod = Y
End If
'***Penn***'
If kludda And Form2.Option1(2).Value = True Then
   
    Pic1.Line (X, Y)-(våg, lod), fargen
     våg = X
    lod = Y
End If
'***Feather***'
If kludda And Form2.Option1(6).Value = True Then Pic1.Line (XX, YY)-(X, Y), Farg
    
'***Stars***'
If kludda And Form2.Option1(8).Value = True Then
    Färgen = Pic1.Point(X, Y)
        For i = 0 To 10
        Randomize
            z = Int(Rnd * -Form1.HScroll2)
            v = Int(Rnd * -Form1.HScroll2)
                Pic1.Line (X - z, Y - v)-(X + z, Y + v), fargen
Next i

    For i = 0 To 10
        Randomize
            a = Int(Rnd * -Form1.HScroll2)
            B = Int(Rnd * -Form1.HScroll2)
                Pic1.Line (X - a, Y + B)-(X + a, Y - B), fargen
Next i
End If
'***air brush***'
If kludda = True And Form2.Option1(9).Value = True Then
    Färgen = Farg
        For i = 0 To Form1.Slider4
            Randomize
                z = Int(Rnd * Form1.Slider3)
                v = Int(Rnd * Form1.Slider3)
                    Pic1.Circle (X + z, Y + v), Form1.S1, Farg
                    Pic1.Circle (X - z, Y - v), Form1.S1, Farg
                    Pic1.Circle (X - z, Y + v), Form1.S1, Farg
                    Pic1.Circle (X + z, Y - v), Form1.S1, Farg
                    
                    
Next i
End If
'If kludda Then Pic1.FillColor = vbRed

'***Rektangle***'
If kludda = True And Form2.Option1(4).Value = True Then
  Färgen = Pic1.Point(X, Y)
        XX2 = X: YY2 = Y
        Shape4.Left = IIf(X > XX, XX, X)
        Shape4.Top = IIf(Y > YY, YY, Y)
        Shape4.Width = Abs(X - XX)
        Shape4.Height = Abs(Y - YY)
End If

'***cirkel***'
If kludda = True And Form2.Option1(0).Value = True Then
    Färgen = Pic1.Point(X, Y)
        XX2 = X: YY2 = Y
        Shape3.Left = IIf(X > XX, XX, X)
        Shape3.Top = IIf(Y > YY, YY, Y)
        Shape3.Width = Abs(X - XX)
        Shape3.Height = Abs(Y - YY)
End If

'***Linje***'
'If kludda = True And Form2.Option1(7).Value = True Then
'    Färgen = Pic1.Point(X, Y)
'    Line2.X2 = X: Line2.Y2 = Y
'    Pic1.DrawWidth = 1
'End If

End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***Rektangle***'
On Error Resume Next
klick = False
If Form2.Option1(4).Value = True Then
    Pic1.DrawWidth = Form1.Slider2.Value
    If XX2 <> XX Then Pic1.Line ((XX), (YY))-(XX2, YY2), fargen, B
        Form1.Shape1.Visible = False
        Shape4.Visible = False
            
            Pic1.Line (XX, YY)-(XX2, YY2), fargen, B
                Pic1.DrawWidth = 1
                upp = 1
                For radie2 = YY To YY2
                upp = upp + 1
                Pic1.Line (XX, YY)-(XX2, radie2), RGB(9, upp, qw), B
                Next
                
                If YY > YY2 Then
                For radie2 = YY2 To YY
                upp = upp + 1
                Pic1.Line (XX, YY)-(XX2, radie2), RGB(1, upp, qw), B
                
                Next
                End If
End If
    
    
'***oval***'
If Form2.Option1(0).Value = True Then
    Pic1.DrawWidth = Form1.s2.Value
        rad = IIf(Abs(YY2 - YY) > Abs(XX2 - XX), Abs(YY2 - YY) / 2, Abs(XX2 - XX) / 2)
    If XX2 <> XX Then Pic1.Circle ((XX2 + XX) / 2, (YY2 + YY) / 2), rad, Farg, , , Abs(YY2 - YY) / Abs(XX2 - XX)
        Shape3.Visible = False
            Pic1.DrawWidth = 1
    radie = rad
    mittfarg = rad
    Pic1.DrawWidth = 3
    For radie = 1 To rad
        mittfarg = mittfarg - 1
        Pic1.Circle ((XX2 + XX) / 2, (YY2 + YY) / 2), radie, RGB(9, mittfarg * 2, qw), , , Abs(YY2 - YY) / Abs(XX2 - XX)
    Next
Pic1.DrawWidth = 1

End If
'***Line***'
'If Form2.Option1(7).Value = True Then
'        Pic1.DrawWidth = Form1.Slider1.Value
' '       Pic1.Line (Line2.X1, Line2.Y1)-(Line2.X2, Line2.Y2), Farg
 '       Pic1.DrawWidth = 1
 '       Line2.Visible = False
'End If
   
    kludda = False
'våg = X
'lod = Y


End Sub
