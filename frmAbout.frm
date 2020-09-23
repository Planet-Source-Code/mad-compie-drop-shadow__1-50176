VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim MouseDown As Boolean
  Dim WithEvents Buttons1 As clsMenuButtons
Attribute Buttons1.VB_VarHelpID = -1

Private Sub Buttons1_Click(ByVal Button As Long, ByVal Index As Long)
  Select Case Index
    Case 0: 'EXIT
            Unload Me
  End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Buttons1.CheckMouseMessage MOUSE_MOVE, Button, x, y
   
  If (y <= 30) And (x < Me.ScaleWidth - 45) Then
    'Menubalk
    If Not (MouseDown) And (Button = vbLeftButton) Then
      MouseDown = True
      ReleaseCapture
      'Send a 'left mouse button down on caption'-message to our form
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
  Else
    If (Button = vbLeftButton) Then MouseDown = True
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Buttons1.CheckMouseMessage MOUSE_DOWN, Button, x, y
  MouseDown = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Buttons1.CheckMouseMessage MOUSE_UP, Button, x, y
  MouseDown = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Buttons1 = Nothing
  Set frmAbout = Nothing
End Sub

Private Sub Form_Load()
  Dim R As Long

  MouseDown = False
  Me.AutoRedraw = True
  Me.Width = 150 * Screen.TwipsPerPixelX
  Me.Height = 150 * Screen.TwipsPerPixelY
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  
  Form_Skin Me, "FILL2"
  
  Me.ForeColor = &H0&
  
  DrawDSText Me.hDC, "Info", 30, 14, vbWhite, vbBlack, FXDS_ShadowDepth3, "Arial", 12, True
  DrawDSText Me.hDC, "ICT builder", 15, 50, &H800000, vbRed, FXDS_ShadowDepth5, "Arial", 12, True
  DrawDSText Me.hDC, "versie " & App.Major & "." & App.Minor, 15, 70, &H71C8FE, &H71C8FE, FXDS_ShadowDepth5, "Arial", 12, True
    
  'Bovenste buttons:
  R = Me.ScaleWidth - 26
  Set Buttons1 = New clsMenuButtons
  Buttons1.Owner = Me
  Buttons1.Add R - 21, 11, R, 32, , , , FXAT_AddTransparent, MOUSE_BTN_LIGHTEN8, MOUSE_BTN_RESAMPLE80
  Buttons1.CreateBackBuffers 21, 21
  Buttons1.Draw 0, 0, LoadResPicture("CLOSE", vbResBitmap)
  Buttons1.Enabled = True
End Sub

