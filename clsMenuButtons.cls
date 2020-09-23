VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
  Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
  Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
  Private Declare Function DrawResIcon Lib "user32" Alias "DrawIcon" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
  
  Public Enum MouseEvents
    MOUSE_MOVE
    MOUSE_UP
    MOUSE_DOWN
  End Enum
  
  Public Enum MouseHoverEffects
    MOUSE_BTN_DARKEN
    MOUSE_BTN_LIGHTEN4
    MOUSE_BTN_LIGHTEN8
    MOUSE_BTN_LIGHTEN20
  End Enum
  Public Enum MouseDownEffects
    MOUSE_BTN_RESAMPLE90
    MOUSE_BTN_RESAMPLE80
  End Enum

  Private Type ButtonType
    ButtonRgn    As Long                      'Physical region of the button
    ButtonsXY    As RECT                      'Client-area of the button
    ButtonWidth  As Long                      'Button width
    ButtonHeight As Long                      'Button height
    BufferX      As Long                      'X-position of button from backbuffer
    Transparent  As FX_AddImage_Transparency  'Button contains transparency color?
    HoverFX      As MouseHoverEffects         'Darken/lighten FX for 'hover' events
    DownFX       As MouseDownEffects          'FX for 'down' events
    Icon         As Variant                   'Icon reference from embedded RES file
    IconSize     As Integer                   'Icon width and height
  End Type
  
  Private cDIB_up        As New clsDIBSection '"up" backbuffer
  Private cDIB_down      As New clsDIBSection '"down" backbuffer
  Private cDIB_over      As New clsDIBSection '"over" backbuffer
  
  Private Buttons()      As ButtonType 'Buttons array
  Private LastButton     As Integer    'Index of button that received last mouse event
  Private ButtonDown     As Integer    'Index of button that received "MouseDown" event
  Private ButtonType     As Integer    '0=none, 1=left, 2=right
  Private ButtonsEnabled As Boolean    'Enable/disable the button FX
  Private Max            As Integer    'UBound index of the Buttons array
  Private Canvas         As Object     'Owner hWnd of the buttons array
  Private BackBuffersOK  As Boolean    'Backbuffers already been created?
  
  Public Event Click(ByVal Button As Long, ByVal Index As Long)
  Public Event MouseOver(ByVal Index As Long)
  Public Event MouseLeave(ByVal Index As Long)

Private Sub Class_Initialize()
  'When the object is created:
  Max = -1
  LastButton = -1
  ButtonDown = -1
  ButtonType = -1
  BackBuffersOK = False
  ButtonsEnabled = False
End Sub

Private Sub Class_Terminate()
  'When the object is deleted:
  Dim i As Integer
  
  For i = 0 To Max
    DeleteObject Buttons(i).ButtonRgn
  Next i
  Erase Buttons
  'cDIB_up , cDIB_down And cDIB_over are automatically cleared up...
End Sub

Public Property Let Owner(obj As Object)
  'Set the owner hWnd for the buttons array
  If (Canvas Is Nothing) Then
    If (TypeOf obj Is Form) Or (TypeOf obj Is PictureBox) Then
      Set Canvas = obj
    Else
      Set Canvas = Nothing
    End If
  End If
End Property

Public Property Get Count() As Integer
  'Get the total number of added buttons
  Count = Max + 1
End Property

Public Property Get Enabled() As Boolean
  Enabled = ButtonsEnabled
End Property

Public Property Let Enabled(How As Boolean)
  ButtonsEnabled = How
End Property

Public Sub Reset()
  If Not (BackBuffersOK) Then Exit Sub
  
  If (LastButton >= 0) Then DrawButtonUp LastButton
  
  LastButton = 0
  ButtonDown = -1
  ButtonType = -1
End Sub

Public Function CreateBackBuffers(ByVal Width As Long, ByVal Height As Long) As Boolean
  'Create 3 backbuffers ("up", "down" and "over") with maximum dimensions:
  '* Width  = max. width for all buttons
  '* Height = max. height for all buttons
    
  CreateBackBuffers = False
    
  If (Canvas Is Nothing) Then Exit Function 'Owner hWnd not yet set!
  If (BackBuffersOK) Then Exit Function     'Already created!
  
  If Not (cDIB_up.Create(Width, Height)) Then Exit Function   'Create the "up" backbuffer
  If Not (cDIB_down.Create(Width, Height)) Then Exit Function 'Create the "down" backbuffer
  If Not (cDIB_over.Create(Width, Height)) Then Exit Function 'Create the "over" backbuffer
  
  BackBuffersOK = True
  CreateBackBuffers = True
End Function

Public Sub Add(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
               Optional ByVal Icon As Variant = vbNullString, _
               Optional ByVal IconSize As Integer = 0, _
               Optional ByVal Rounded As Boolean = False, _
               Optional ByVal Transparent As FX_AddImage_Transparency = FXAT_AddNormal, _
               Optional ByVal HoverFX As MouseHoverEffects = MOUSE_BTN_LIGHTEN8, _
               Optional ByVal DownFX As MouseDownEffects = MOUSE_BTN_RESAMPLE80)
  'Left, Top, Right, Bottom = RECT dimensions
  'Icon = offset of icon in RES file ("" is no icon)
  'IconSize = icon width
  'Rounded = create elliptical region?
  'Transparent = how to add the button?
  'HoverFX = which FX when hover event occurred?
  'DownFX = which FX when down event occurred?
  If (Canvas Is Nothing) Then Exit Sub

  Max = Max + 1
  ReDim Preserve Buttons(0 To Max) As ButtonType
  Buttons(Max).Icon = Icon
  Buttons(Max).IconSize = IconSize
  SetRect Buttons(Max).ButtonsXY, Left, Top, Right, Bottom
  Buttons(Max).ButtonWidth = Right - Left
  Buttons(Max).ButtonHeight = Bottom - Top
  If (Max = 0) Then
    Buttons(Max).BufferX = 0 'Eerste knop...
  Else
    Buttons(Max).BufferX = Buttons(Max - 1).BufferX + Buttons(Max - 1).ButtonWidth
  End If
  If (Rounded) Then
    Buttons(Max).ButtonRgn = CreateEllipticRgnIndirect(Buttons(Max).ButtonsXY)
  Else
    Buttons(Max).ButtonRgn = CreateRectRgn(Left, Top, Right, Bottom)
  End If
  Buttons(Max).Transparent = Transparent
  Buttons(Max).HoverFX = HoverFX
  Buttons(Max).DownFX = DownFX
End Sub

Private Sub DrawButtonUp(Index As Integer)
  'Draw & refresh the button "up" from the backbuffer "up"
  cDIB_up.Object2DC Canvas.hDC, Buttons(Index).ButtonsXY.Left, Buttons(Index).ButtonsXY.Top, Buttons(Index).ButtonWidth, Buttons(Index).ButtonHeight, Buttons(Index).BufferX, 0
  RedrawWindow Canvas.hWnd, Buttons(Index).ButtonsXY, 0&, &H1
End Sub

Private Sub DrawButtonOver(Index As Integer)
  'Draw & refresh the button "over" from the backbuffer "over"
  cDIB_over.Object2DC Canvas.hDC, Buttons(Index).ButtonsXY.Left, Buttons(Index).ButtonsXY.Top, Buttons(Index).ButtonWidth, Buttons(Index).ButtonHeight, Buttons(Index).BufferX, 0
  RedrawWindow Canvas.hWnd, Buttons(Index).ButtonsXY, 0&, &H1
End Sub

Private Sub DrawButtonDown(Index As Integer)
  'Draw & refresh the button "down" from the backbuffer "down"
  cDIB_down.Object2DC Canvas.hDC, Buttons(Index).ButtonsXY.Left, Buttons(Index).ButtonsXY.Top, Buttons(Index).ButtonWidth, Buttons(Index).ButtonHeight, Buttons(Index).BufferX, 0
  RedrawWindow Canvas.hWnd, Buttons(Index).ButtonsXY, 0&, &H1
End Sub

Public Sub Draw(ByVal ButtonFrom As Integer, ButtonTo As Integer, ByRef P As StdPicture)
  'Draw the buttons from the button array [from..to] to the canvas
  'ButtonFrom is LBound button index
  'ButtonTo   is UBound button index
  'P contains the original button to draw
  Dim i           As Integer
  Dim W           As Long
  Dim H           As Long
  Dim XY          As POINTAPI
  Dim cDIB_backgr As New clsDIBSection
  Dim cDIB_mem    As New clsDIBSection
    
  If (Canvas Is Nothing) Then Exit Sub                    'Canvas not yet been set!
  If (ButtonFrom < 0) Or (ButtonFrom > Max) Then Exit Sub 'Illegal index!
  If (ButtonTo < 0) Or (ButtonTo > Max) Then Exit Sub     'Illegal index!
  
  'Make memory DIB for the actual background image in "backgr":
  cDIB_backgr.Create Buttons(ButtonFrom).ButtonWidth, Buttons(ButtonFrom).ButtonHeight
  
  'Get & test the button's dimensions:
  cDIB_backgr.GetPictureWH P, W, H
  If (W > Buttons(ButtonFrom).ButtonWidth) Then Exit Sub   'This picture's width is too large!
  If (H > Buttons(ButtonFrom).ButtonHeight) Then Exit Sub  'This picture's height is too large!
  
  'Draw the buttons:
  For i = ButtonFrom To ButtonTo
    'Save the original button as "mem":
     cDIB_mem.CreateFromPicture P
    'Draw icon (if present) in "mem":
     If (Buttons(i).Icon <> "") Then
       On Error Resume Next
       DrawResIcon cDIB_mem.hDC, (W - Buttons(i).IconSize) / 2, (H - Buttons(i).IconSize) / 2, LoadResPicture(Buttons(i).Icon, vbResIcon)
       On Error GoTo 0
     End If
    '********* UP ********************************************
     'Put background from Canvas to "backgr":
      cDIB_backgr.DC2Object Canvas.hDC, Buttons(i).ButtonsXY.Left, Buttons(i).ButtonsXY.Top, W, H, vbSrcCopy
     'Button & icon from "mem" to "backgr" and then to "up":
      cDIB_backgr.FX_AddImage cDIB_mem, 0, 0, Buttons(i).Transparent
      cDIB_backgr.Object2DC cDIB_up.hDC, Buttons(i).BufferX, 0, W, H, 0, 0, vbSrcCopy
    '********* OVER ******************************************
     'Put background from Canvas to "backgr":
      cDIB_backgr.DC2Object Canvas.hDC, Buttons(i).ButtonsXY.Left, Buttons(i).ButtonsXY.Top, W, H, vbSrcCopy
     'Lighten/darken the button & icon from "mem":
      Select Case Buttons(i).HoverFX
        Case MOUSE_BTN_LIGHTEN4:  cDIB_mem.FX_Lighten 10
        Case MOUSE_BTN_LIGHTEN8:  cDIB_mem.FX_Lighten 20
        Case MOUSE_BTN_LIGHTEN20: cDIB_mem.FX_Lighten 51
        Case MOUSE_BTN_DARKEN:    cDIB_mem.FX_Darken 200
      End Select
     'Button & icon from "mem" to "backgr" and then to "over":
      cDIB_backgr.FX_AddImage cDIB_mem, 0, 0, Buttons(i).Transparent
      cDIB_backgr.Object2DC cDIB_over.hDC, Buttons(i).BufferX, 0, W, H, 0, 0, vbSrcCopy
    '********* DOWN ******************************************
     'Put background from Canvas to "backgr":
      cDIB_backgr.DC2Object Canvas.hDC, Buttons(i).ButtonsXY.Left, Buttons(i).ButtonsXY.Top, W, H, vbSrcCopy
      Select Case Buttons(i).DownFX
        Case MOUSE_BTN_RESAMPLE80: '(Resample to 80%)
                                   cDIB_mem.FX_Resample W * 0.8, H * 0.8
                                   XY.X = (W - W * 0.8) \ 2: XY.Y = (H - H * 0.8) \ 2
        Case MOUSE_BTN_RESAMPLE90: '(Resample to 90%)
                                   cDIB_mem.FX_Resample W * 0.9, H * 0.9
                                   XY.X = (W - W * 0.9) \ 2: XY.Y = (H - H * 0.9) \ 2
      End Select
     'Button & icon from "mem" to "backgr" and then to "down":
      cDIB_backgr.FX_AddImage cDIB_mem, XY.X, XY.Y, Buttons(i).Transparent
      cDIB_backgr.Object2DC cDIB_down.hDC, Buttons(i).BufferX, 0, W, H, 0, 0, vbSrcCopy
    '*********************************************************
    'Always draw the button "up":
     DrawButtonUp i
    Next i
  End Sub

Public Sub CheckMouseMessage(ByVal Message As MouseEvents, ByVal Button As Integer, ByVal X As Long, ByVal Y As Long)
  Dim i         As Integer
  Dim NewButton As Integer
  Dim TmpButton As Integer
    
  If Not (ButtonsEnabled) Then Exit Sub
  
  NewButton = -1
  TmpButton = -1
  
  If (Message = MOUSE_DOWN) Then ButtonType = Button And &H3 'Left(1) or right(2)
        
  For i = 0 To Max
    If PtInRegion(Buttons(i).ButtonRgn, X, Y) Then
      If (ButtonDown < 0) Then NewButton = i
      TmpButton = i
      Exit For
    End If
  Next i
  
  If (LastButton >= 0) Then
    If (NewButton <> LastButton) Then
      If (ButtonDown < 0) Then
        DrawButtonUp LastButton
        RaiseEvent MouseLeave(LastButton)
        LastButton = -1
      End If
    End If
  End If
    
  Select Case Message
    Case MOUSE_MOVE:
                        If (NewButton >= 0) Then
                          If (LastButton <> NewButton) Then
                            DrawButtonOver NewButton
                            RaiseEvent MouseOver(NewButton)
                          End If
                          LastButton = NewButton
                        End If
    Case MOUSE_DOWN:
                        If (NewButton >= 0) Then
                          DrawButtonDown NewButton
                          LastButton = NewButton
                          ButtonDown = NewButton
                        End If
    Case MOUSE_UP:
                        If (ButtonDown >= 0) Then
                          LastButton = ButtonDown
                          If (TmpButton = ButtonDown) Then
                            DrawButtonOver ButtonDown
                            RaiseEvent Click(ButtonType, ButtonDown)
                          Else
                            DrawButtonUp ButtonDown
                          End If
                        End If
                        ButtonDown = -1
  End Select
End Sub

