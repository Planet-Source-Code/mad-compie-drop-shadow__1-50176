Attribute VB_Name = "modGFX"
Option Explicit

Private Type LOGFONT
  lfHeight         As Long
  lfWidth          As Long
  lfEscapement     As Long
  lfOrientation    As Long
  lfWeight         As Long
  lfItalic         As Byte
  lfUnderline      As Byte
  lfStrikeOut      As Byte
  lfCharSet        As Byte
  lfOutPrecision   As Byte
  lfClipPrecision  As Byte
  lfQuality        As Byte
  lfPitchAndFamily As Byte
  lfFaceName       As String * 32
End Type

Public Type RGBQUAD
  rgbBlue     As Byte
  rgbGreen    As Byte
  rgbRed      As Byte
  rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER '40 bytes
  biSize          As Long
  biWidth         As Long
  biHeight        As Long
  biPlanes        As Integer
  biBitCount      As Integer
  biCompression   As Long
  biSizeImage     As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed       As Long
  biClrImportant  As Long
End Type

Public Type BITMAPINFO
  bmiHeader       As BITMAPINFOHEADER
  bmiColors       As RGBQUAD
End Type

Private Type DWORD
  Low  As Integer
  High As Integer
End Type

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

' Bitmap Header Definition
Private Type BITMAP '14 bytes
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Sub CopyMemoryLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Private Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Public Enum FX_AddImage_Transparency
  FXAT_AddTransparent = -1
  FXAT_AddNormal = 0
  FXAT_AddDarkest = 1
  FXAT_AddAlphaBlended = 2
  FXAT_Combine = 3
  FXAT_CombineTransparent = 4
End Enum
  
Public Enum FX_DropShadow_ShadowDepth
  FXDS_ShadowDepth1 = 1
  FXDS_ShadowDepth3 = 3
  FXDS_ShadowDepth5 = 5
  FXDS_ShadowDepth7 = 7
  FXDS_ShadowDepth9 = 9
End Enum
  
Public Enum FX_FlipType
  FX_Flip_Vertical
  FX_Flip_Horizontal
End Enum

Private Type TRIVERTEX
  x     As Long
  y     As Long
  Red   As Integer 'Ushort value
  Green As Integer 'Ushort value
  Blue  As Integer 'ushort value
  Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
  UpperLeft  As Long  'In reality this is a UNSIGNED Long
  LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
  'Convert automation color to Windows color
  If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = -1 'CLR_INVALID
  End If
End Function

Private Sub SetTriVertexColor(tTV As TRIVERTEX, lColor As Long)
  Dim lRed   As Long
  Dim lGreen As Long
  Dim lBlue  As Long
  
  lRed = (lColor And &HFF&) * &H100&
  lGreen = (lColor And &HFF00&)
  lBlue = (lColor And &HFF0000) \ &H100&
  SetTriVertexColorComponent tTV.Red, lRed
  SetTriVertexColorComponent tTV.Green, lGreen
  SetTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub SetTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
  If ((lComponent And &H8000&) = &H8000&) Then
    iColor = (lComponent And &H7F00&)
    iColor = iColor Or &H8000
  Else
    iColor = lComponent
  End If
End Sub

Public Sub DrawDSText(hDC As Long, Text As String, Xpos As Long, Ypos As Long, FontColor As Long, ShadowColor As Long, ShadowDepth As FX_DropShadow_ShadowDepth, FontName As String, FontSize As Long, FontBold As Boolean, Optional Centered As Boolean = False)
  Dim L         As Long
  Dim H         As Long
  Dim W         As Long
  Dim Font      As LOGFONT
  Dim hFont     As Long
  Dim oFont     As Long
  Dim DIBresult As clsDIBSection
  Dim DIBtext   As clsDIBSection
  Dim DIBshadow As clsDIBSection
      
  'Create Font for DC:
  With Font
    .lfHeight = -(FontSize * 20) / Screen.TwipsPerPixelY ' set font size
    .lfFaceName = FontName & Chr(0) 'apply font name
    If (FontBold) Then
      .lfWeight = 700 'Bold
    Else
      .lfWeight = 0   'Normal
    End If
  End With
  hFont = CreateFontIndirect(Font)
   
  'Get text extents:
  oFont = SelectObject(hDC, hFont)
  L = GetTabbedTextExtent(hDC, Text, Len(Text), 0, 0)
  H = (L \ &H10000) + 10   'Is +5 pixels top & bottom
  W = (L And &HFFFF&) + 10 'Is +5 pixels left & right
  SelectObject hDC, oFont
  
  'Create DIB sections:
  Set DIBresult = New clsDIBSection 'Background and text
  Set DIBtext = New clsDIBSection   'Only AA Text
  Set DIBshadow = New clsDIBSection 'Only AA Text shadowed
  If Not (DIBresult.Create(W, H)) Then GoTo DIBerror
  If Not (DIBtext.Create(W, H)) Then GoTo DIBerror
  If Not (DIBshadow.Create(W, H)) Then GoTo DIBerror
  
  'Put current background into DIBresult
  DIBresult.DC2Object hDC, Xpos - 5, Ypos - 5, W, H, vbSrcCopy
  'Draw AA text on blank DIB:
  oFont = SelectObject(DIBtext.hDC, hFont)
  SetTextColor DIBtext.hDC, ShadowColor
  TabbedTextOut DIBtext.hDC, 5, 5, Text, Len(Text), 0, 0, 0
  SelectObject DIBtext.hDC, oFont
  'Copy AA text to DIBshadow:
  DIBshadow.DC2Object DIBtext.hDC, 0, 0
  'Perform drop shadow effect, result in DIBshadow:
  DIBtext.FX_DropShadow DIBshadow, ShadowDepth
  DIBresult.FX_AddImage DIBshadow, 0, 0, FXAT_AddAlphaBlended
  'Draw AA text again on (textured) background:
  oFont = SelectObject(DIBresult.hDC, hFont)
  SetTextColor DIBresult.hDC, FontColor
  TabbedTextOut DIBresult.hDC, 4, 4, Text, Len(Text), 0, 0, 0
  SelectObject DIBresult.hDC, oFont
  'Draw result back to destination DC:
  DIBresult.Object2DC hDC, Xpos - 5, Ypos - 5, W, H, 0, 0, vbSrcCopy

DIBerror:
  'Clean up DIB sections:
  Set DIBshadow = Nothing
  Set DIBtext = Nothing
  Set DIBresult = Nothing
  'Clean up user font:
  DeleteObject hFont
End Sub

Public Sub Draw3DText(hDC As Long, Text As String, Xpos As Long, Ypos As Long, ForeColor As Long, ShadowColor As Long, FontName As String, FontSize As Long)
  Dim L     As Long
  Dim Color As Long
  Dim Font  As LOGFONT
  Dim hFont As Long
    
  With Font
    .lfHeight = -(FontSize * 20) / Screen.TwipsPerPixelY ' set font size
    .lfFaceName = FontName & Chr(0) 'apply font name
    .lfWeight = 700   'this is how bold the font is .. apply a in param if you want
  End With
    
  hFont = CreateFontIndirect(Font)
  SelectObject hDC, hFont
  L = GetTabbedTextExtent(hDC, Text, Len(Text), 0, 0)
  Color = GetTextColor(hDC)
  SetTextColor hDC, ShadowColor
  TabbedTextOut hDC, Xpos + 1, Ypos + 1, Text, Len(Text), 0, 0, 0
  SetTextColor hDC, ForeColor
  TabbedTextOut hDC, Xpos, Ypos, Text, Len(Text), 0, 0, 0
  SetTextColor hDC, Color
  DeleteObject hFont
End Sub

Public Sub Form_Skin(F As Form, Fill As String)
  Dim hFRgn      As Long
  Dim FillC      As Long
  Dim W          As Long
  Dim H          As Long
  Dim StartColor As Long
  Dim EndColor   As Long
  Dim vert(1)    As TRIVERTEX
  Dim gRect      As GRADIENT_RECT

  W = F.ScaleWidth - 1
  H = F.ScaleHeight - 1
       
  TileBlt F.hwnd, F.hDC, LoadResPicture(Fill, vbResBitmap)
  
  StartColor = TranslateColor(&H71C8FE)
  EndColor = TranslateColor(&HCEE9FA)
  SetTriVertexColor vert(0), StartColor
  vert(0).x = 0
  vert(0).y = 0
  SetTriVertexColor vert(1), EndColor
  vert(1).x = W
  vert(1).y = 35
  gRect.UpperLeft = 0
  gRect.LowerRight = 1
  GradientFillRect F.hDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V

  
  '4 transparent corners:
  TransBlt F.hDC, 0, 0, 30, 30, LoadResPicture("CORNERUL", vbResBitmap), 0, 0, &HFFFFFF
  TransBlt F.hDC, W - 30, 0, 30, 30, LoadResPicture("CORNERUR", vbResBitmap), 0, 0, &HFFFFFF
  TransBlt F.hDC, 0, H - 30, 30, 30, LoadResPicture("CORNERLL", vbResBitmap), 0, 0, &HFFFFFF
  TransBlt F.hDC, W - 30, H - 30, 30, 30, LoadResPicture("CORNERLR", vbResBitmap), 0, 0, &HFFFFFF
  
  'Line left:
  F.PaintPicture LoadResPicture("LINELEFT", vbResBitmap), 0, 30, , H - 60
  'Line right:
  F.PaintPicture LoadResPicture("LINERIGHT", vbResBitmap), W - 10, 30, , H - 60
  'Line upper:
  F.PaintPicture LoadResPicture("LINEUPPER", vbResBitmap), 30, 0, W - 60
  'Line lower:
  F.PaintPicture LoadResPicture("LINELOWER", vbResBitmap), 30, H - 10, W - 60
  
  hFRgn = CreateRoundRectRgn(0, 0, W + 1, H + 1, 50, 50)
  F.ForeColor = &HC0E0FF 'orange
  RoundRect F.hDC, 10, 10, W - 10 + 1, H - 10 + 1, 40, 40
  SetWindowRgn F.hwnd, hFRgn, True
    
  DeleteObject hFRgn
End Sub

Public Sub TileBlt(ByVal hWndDest As Long, ByVal hDCDest As Long, ByVal hBmpSrc As Long)
   'hWndDest:     Destination hWnd of Form, PictureBox, ...
   'hDCDest:      Destination hDC of Form, PictureBox, ...
   'hBmpSrc:      hBitmap of source (Picture property)
   '
   Dim bmp     As BITMAP  ' Header info for passed bitmap handle
   Dim hDCSrc  As Long    ' Device context for source
   Dim hBmpTmp As Long    ' Holding space for temporary bitmap
   Dim dRect   As RECT    ' Holds coordinates of destination rectangle
   Dim Rows    As Long    ' Number of rows in destination
   Dim Cols    As Long    ' Number of columns in destination
   Dim dX      As Long    ' CurrentX in destination
   Dim dy      As Long    ' CurrentY in destination
   Dim i       As Long
   Dim j       As Long
   
   ' Get destination rectangle and device context.
   GetClientRect hWndDest, dRect
   
   'Create source DC and select passed bitmap into it.
   hDCSrc = CreateCompatibleDC(hDCDest)
   hBmpTmp = SelectObject(hDCSrc, hBmpSrc)
   
   'Get size information about passed bitmap, and
   'Calc number of rows and columns to paint.
   GetObject hBmpSrc, Len(bmp), bmp
   Rows = dRect.Right \ bmp.bmWidth
   Cols = dRect.Bottom \ bmp.bmHeight
   
   'Spray out across destination.
   For i = 0 To Rows
     dX = i * bmp.bmWidth
     For j = 0 To Cols
       dy = j * bmp.bmHeight
       BitBlt hDCDest, dX, dy, bmp.bmWidth, bmp.bmHeight, hDCSrc, 0, 0, vbSrcCopy
     Next j
   Next i
   
   'and clean up...
   SelectObject hDCSrc, hBmpTmp
   DeleteDC hDCSrc
End Sub

Public Sub TransBlt(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hBmpSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal TransColor As Long)
   'hDestDC:     Destination device context
   'x, y:        Upper-left destination coordinates (pixels)
   'nWidth:      Width of destination
   'nHeight:     Height of destination
   'hBmpSrc:     hBitmap of source (Picture property)
   'xSrc, ySrc:  Upper-left source coordinates (pixels)
   'TransColor:  RGB value for transparent pixels
   '
   Dim OrigColor    As Long   'Holds original background color
   Dim hDCSrc       As Long   'Device context for source
   Dim hBmpTmp      As Long   'Holding space for temporary bitmap
   Dim saveDC       As Long   'Backup copy of source bitmap
   Dim maskDC       As Long   'Mask bitmap (monochrome)
   Dim invDC        As Long   'Inverse of mask bitmap (monochrome)
   Dim resultDC     As Long   'Combination of source bitmap & background
   Dim hSaveBmp     As Long   'Bitmap stores backup copy of source bitmap
   Dim hMaskBmp     As Long   'Bitmap stores mask (monochrome)
   Dim hInvBmp      As Long   'Bitmap holds inverse of mask (monochrome)
   Dim hResultBmp   As Long   'Bitmap combination of source & background
   Dim hSavePrevBmp As Long   'Holds previous bitmap in saved DC
   Dim hMaskPrevBmp As Long   'Holds previous bitmap in the mask DC
   Dim hInvPrevBmp  As Long   'Holds previous bitmap in inverted mask DC
   Dim hDestPrevBmp As Long   'Holds previous bitmap in destination DC
      
   'Create source DC and select passed bitmap into it:
   hDCSrc = CreateCompatibleDC(hDestDC)
   hBmpTmp = SelectObject(hDCSrc, hBmpSrc)
            
   'Create DCs to hold various stages of transformation:
   saveDC = CreateCompatibleDC(hDestDC)
   maskDC = CreateCompatibleDC(hDestDC)
   invDC = CreateCompatibleDC(hDestDC)
   resultDC = CreateCompatibleDC(hDestDC)
      
   'Create monochrome bitmaps for the mask-related bitmaps:
   hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
   hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
      
   'Create color bitmaps for final result & stored copy of source:
   hResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
   hSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
      
   'Select bitmaps into DCs:
   hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
   hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
   hInvPrevBmp = SelectObject(invDC, hInvBmp)
   hDestPrevBmp = SelectObject(resultDC, hResultBmp)
      
   'Make backup of source bitmap to restore later:
   BitBlt saveDC, 0, 0, nWidth, nHeight, hDCSrc, xSrc, ySrc, vbSrcCopy
   
   'Create mask: set background color of source to transparent color:
   OrigColor = SetBkColor(hDCSrc, TransColor)
   BitBlt maskDC, 0, 0, nWidth, nHeight, hDCSrc, xSrc, ySrc, vbSrcCopy
   TransColor = SetBkColor(hDCSrc, OrigColor)
      
   'Create inverse of mask to AND w/ source & combine w/ background:
   BitBlt invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy
      
   'Copy background bitmap to result & create final transparent bitmap:
   BitBlt resultDC, 0, 0, nWidth, nHeight, hDestDC, x, y, vbSrcCopy
      
   'AND mask bitmap w/ result DC to punch hole in the background by
   'painting black area for non-transparent portion of source bitmap:
   BitBlt resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd
     
   'AND inverse mask w/ source bitmap to turn off bits associated
   'with transparent area of source bitmap by making it black:
   BitBlt hDCSrc, xSrc, ySrc, nWidth, nHeight, invDC, 0, 0, vbSrcAnd
      
   'XOR result w/ source bitmap to make background show through:
   BitBlt resultDC, 0, 0, nWidth, nHeight, hDCSrc, xSrc, ySrc, vbSrcPaint
      
   'Display transparent bitmap on background:
   BitBlt hDestDC, x, y, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy
      
   'Restore backup of original bitmap:
   BitBlt hDCSrc, xSrc, ySrc, nWidth, nHeight, saveDC, 0, 0, vbSrcCopy
      
   'Select original objects back:
   SelectObject saveDC, hSavePrevBmp
   SelectObject resultDC, hDestPrevBmp
   SelectObject maskDC, hMaskPrevBmp
   SelectObject invDC, hInvPrevBmp
      
   'Deallocate system resources:
   DeleteObject (hSaveBmp)
   DeleteObject (hMaskBmp)
   DeleteObject (hInvBmp)
   DeleteObject (hResultBmp)
   DeleteDC (saveDC)
   DeleteDC (invDC)
   DeleteDC (maskDC)
   DeleteDC (resultDC)
      
   'Deallocate source bitmap memory:
   SelectObject hDCSrc, hBmpTmp
   DeleteDC hDCSrc
End Sub

