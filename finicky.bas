Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function ExtFloodFill Lib "Gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2


#If Win32 Then
Public Const HWND_TOPMOST& = -1
#Else
Public Const HWND_TOPMOST& = -1
#End If

#If Win32 Then
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#Else
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#End If
#If Win32 Then
 Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
 Declare Sub SetWindowPos Lib "user" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If
