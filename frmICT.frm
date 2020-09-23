VERSION 5.00
Begin VB.Form frmICT 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmICT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmICT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim MouseDown As Boolean
  
  Private WithEvents Buttons1 As clsMenuButtons
Attribute Buttons1.VB_VarHelpID = -1
  Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1
  
Private Sub ShowBalloon()
   m_frmSysTray.ShowBalloonTip "Welkom bij ICT", "ICT", NIIF_INFO
End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
  Select Case sKey
    Case "open":  Me.Show
                  Me.ZOrder
    Case "einde": Unload Me
    Case "info":  frmAbout.Show vbModal
    Case Else:    MsgBox "Clicked item with key " & sKey, vbInformation
  End Select
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
  If (Me.Enabled) Then
    Me.Show
    Me.ZOrder
  End If
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
  If (eButton = vbRightButton) Then m_frmSysTray.ShowMenu
End Sub

Private Sub Buttons1_Click(ByVal Button As Long, ByVal Index As Long)
  Select Case Index
    Case 0: 'EXIT
            Unload Me
    Case 1: 'MINIMIZE
            Me.Hide
    Case 2: 'HELP
            frmAbout.Show vbModal
  End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Check for menubar button events:
  Buttons1.CheckMouseMessage MOUSE_MOVE, Button, x, y
   
  If (y <= 30) And (x < Me.ScaleWidth - 93) Then
    'Move the form?
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload m_frmSysTray
  Set m_frmSysTray = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Buttons1 = Nothing
  Set frmICT = Nothing
  End
End Sub

Private Sub Form_Load()
  Dim R As Long
    
  MouseDown = False
  Me.AutoRedraw = True
  Me.Width = 750 * Screen.TwipsPerPixelX
  Me.Height = 500 * Screen.TwipsPerPixelY
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
  Form_Skin Me, "FILL"
    
  DrawDSText Me.hDC, "ICT builder", 30, 14, vbWhite, &H0&, FXDS_ShadowDepth3, "Arial", 12, True
  DrawDSText Me.hDC, "Just a test", 30, 140, vbBlack, &H0&, FXDS_ShadowDepth7, "Arial", 24, True
         
  'Menubar buttons:
  Set Buttons1 = New clsMenuButtons
  Buttons1.Owner = Me
  R = Me.ScaleWidth - 26
  Buttons1.Add R - 21, 11, R, 32, , , , FXAT_AddTransparent, MOUSE_BTN_LIGHTEN8, MOUSE_BTN_RESAMPLE80
  Buttons1.Add R - 21 - 2 - 21, 11, R - 21 - 2, 32, , , , FXAT_AddTransparent, MOUSE_BTN_LIGHTEN8, MOUSE_BTN_RESAMPLE80
  Buttons1.Add R - 21 - 2 - 21 - 2 - 21, 11, R - 21 - 2 - 21 - 2, 32, , , , FXAT_AddTransparent, MOUSE_BTN_LIGHTEN8, MOUSE_BTN_RESAMPLE80
  Buttons1.CreateBackBuffers 63, 21
  Buttons1.Draw 0, 0, LoadResPicture("CLOSE", vbResBitmap)
  Buttons1.Draw 1, 1, LoadResPicture("MINIMIZE", vbResBitmap)
  Buttons1.Draw 2, 2, LoadResPicture("HELP", vbResBitmap)
  Buttons1.Enabled = True
    
  'Put into systray
  Set m_frmSysTray = New frmSysTray
  With m_frmSysTray
    .AddMenuItem "&Open ICT builder", "open", True
    .AddMenuItem "-"
    .AddMenuItem "&vbAccelerator on the Web", "vbAccelerator"
    .AddMenuItem "&Info...", "info"
    .AddMenuItem "-"
    .AddMenuItem "&Einde", "einde"
    .ToolTip = "SysTray Sample!"
  End With
  m_frmSysTray.IconHandle = Me.Icon.Handle
End Sub

