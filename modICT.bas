Attribute VB_Name = "modICT"
'ICT project
'This is a preliminary project for the documentation
'of the ICT means of a company
'* GFX part: OK
'* ...
'* (c)2003 by Mad Compie, Kuurne, Belgium

Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Sub Main()
  frmICT.Show
End Sub
