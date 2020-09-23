VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
  End Type
  
  Private Type SAFEARRAY2D
    cDims          As Integer
    fFeatures      As Integer
    cbElements     As Long
    cLocks         As Long
    pvData         As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
  End Type
  
  Private Declare Function GetDesktopWindow Lib "user32" () As Long
  Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
  Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
  Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
  
  Private Const BI_RGB = 0&
  Private Const BI_RLE4 = 2&
  Private Const BI_RLE8 = 1&
  Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

  Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
  End Type
  
  Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
  
  Private m_hDIb     As Long       'Handle to the current DIBSection
  Private m_hBmpOld  As Long       'Handle to the old bitmap in the DC, for clear up
  Private m_hDC      As Long       'Handle to the Device context holding the DIBSection
  Private m_lPtr     As Long       'Address of memory pointing to the DIBSection's bits
  Private m_tBI      As BITMAPINFO 'Type containing the Bitmap information
  Private m_BC       As Long       'Background color
  Private m_TC       As Long       'Transparent color is pixel at upper-left corner
  
  Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Class_Initialize()
  'Debug.Print "INITIALISING:" & m_hDC
End Sub

Private Sub Class_Terminate()
  'Debug.Print "TERMINATING: " & m_hDC
  ClearUp
End Sub

Public Sub ClearUp()
  If (m_hDC <> 0) Then
    If (m_hDIb <> 0) Then
      SelectObject m_hDC, m_hBmpOld
      DeleteObject m_hDIb
    End If
    DeleteObject m_hDC
  End If
  m_hDC = 0
  m_hDIb = 0
  m_hBmpOld = 0
  m_lPtr = 0
End Sub

Private Function CreateDIB(ByVal lHDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByRef hDib As Long) As Boolean
  With m_tBI.bmiHeader
    .biSize = Len(m_tBI.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine * .biHeight
  End With
  hDib = CreateDIBSection(lHDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
  CreateDIB = (hDib <> 0)
End Function

Public Function CreateFromPicture(ByRef P As StdPicture)
  Dim lHDC        As Long
  Dim lhDCDesktop As Long
  Dim lhBmpOld    As Long
  Dim tBMP        As BITMAP
    
  GetObjectAPI P.Handle, Len(tBMP), tBMP
  If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
    lhDCDesktop = GetDC(GetDesktopWindow())
    If (lhDCDesktop <> 0) Then
      lHDC = CreateCompatibleDC(lhDCDesktop)
      DeleteDC lhDCDesktop
      If (lHDC <> 0) Then
        lhBmpOld = SelectObject(lHDC, P.Handle)
        DC2Object lHDC
        SelectObject lHDC, lhBmpOld
        DeleteObject lHDC
        m_TC = GetPixel(m_hDC, 0, 0) 'Set the transparent colour from (0,0) - can manually be altered!
      End If
    End If
  End If
End Function

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
  ClearUp
  m_hDC = CreateCompatibleDC(0)
  If (m_hDC <> 0) Then
    If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
      m_hBmpOld = SelectObject(m_hDC, m_hDIb)
      BackColor = vbWhite
      Create = True
    Else
      DeleteObject m_hDC
      m_hDC = 0
    End If
  End If
End Function

Public Sub GetPictureWH(ByRef P As StdPicture, ByRef Width As Long, ByRef Height As Long)
  Dim tBMP As BITMAP
    
  GetObjectAPI P.Handle, Len(tBMP), tBMP
  Width = tBMP.bmWidth
  Height = tBMP.bmHeight
End Sub

Public Property Get TransparentColor() As Long
  TransparentColor = m_TC
End Property

Public Property Let TransparentColor(TColour As Long)
  m_TC = TColour
End Property

Public Property Get BytesPerScanLine() As Long
  'Scans must align on dword boundaries:
  BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
  Width = m_tBI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
  Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub DC2Object(ByVal lSrcDC As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
  'From source DC to current DIB object's DC
  If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biWidth
  If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biHeight
  BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lSrcDC, lSrcLeft, lSrcTop, eRop
  'Set transparent colour from (0,0):
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Public Sub Object2DC(ByVal lDestDC As Long, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
  'From current DIB object's DC to destination DC
  If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
  If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
  BitBlt lDestDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hDC() As Long
  hDC = m_hDC
End Property

Public Property Get hDib() As Long
  hDib = m_hDIb
End Property

Public Property Get DIBSectionBitsPtr() As Long
  DIBSectionBitsPtr = m_lPtr
End Property

Public Property Let BackColor(C As Long)
  Dim R   As RECT
  Dim hBr As Long
  
  If (m_hDC = 0) Then Exit Property
  
  hBr = CreateSolidBrush(C)
  SetBkMode m_hDC, 1 'TRANSPARENT=1
  SetRect R, 0, 0, Me.Width, Me.Height
  FillRect m_hDC, R, hBr
  DeleteObject hBr
  m_BC = C
End Property

Public Property Get BackColor() As Long
  BackColor = m_BC
End Property

Public Sub FX_Flip(FlipType As FX_FlipType)
  If (FlipType = FX_Flip_Horizontal) Then
    StretchBlt m_hDC, 0, 0, Me.Width, Me.Height, m_hDC, Me.Width - 1, 0, -Me.Width, Me.Height, vbSrcCopy
  End If
  
  If (FlipType = FX_Flip_Vertical) Then
    StretchBlt m_hDC, 0, 0, Me.Width, Me.Height, m_hDC, 0, Me.Height - 1, Me.Width, -Me.Height, vbSrcCopy
  End If
  
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Public Sub FX_Resample(ByVal lNewWidth As Long, ByVal lNewHeight As Long)
  'Smoothly resample the current DIB object to another DIB object (SMALLER, NOT LARGER!)
  Dim bDibFrom() As Byte
  Dim bDibTo()   As Byte
  Dim tSAFrom    As SAFEARRAY2D
  Dim tSATo      As SAFEARRAY2D
  Dim xScale     As Single
  Dim yScale     As Single
  Dim x          As Long
  Dim y          As Long
  Dim xEnd       As Long
  Dim xOut       As Long
  Dim FX         As Single
  Dim fY         As Single
  Dim ifY        As Long
  Dim ifX        As Long
  Dim dX         As Single
  Dim dy         As Single
  Dim cDibTo     As clsDIBSection
  
  Dim R As Long, R1 As Single, R2 As Single, R3 As Single, R4 As Single
  Dim G As Long, G1 As Single, G2 As Single, G3 As Single, G4 As Single
  Dim b As Long, B1 As Single, B2 As Single, B3 As Single, B4 As Single
  Dim iR1 As Long, iG1 As Long, iB1 As Long
  Dim iR2 As Long, iG2 As Long, iB2 As Long

  If (lNewWidth = Me.Width) And (lNewHeight = Me.Height) Then Exit Sub 'No resampling...
    
  Set cDibTo = New clsDIBSection
  
  If Not cDibTo.Create(lNewWidth, lNewHeight) Then Exit Sub
  
  'Get the bits in the from DIB section:
  With tSAFrom
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = Me.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = Me.BytesPerScanLine()
    .pvData = m_lPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

  'Get the bits in the to DIB section:
  With tSATo
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = cDibTo.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = cDibTo.BytesPerScanLine()
    .pvData = cDibTo.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

  xScale = Me.Width / cDibTo.Width
  yScale = Me.Height / cDibTo.Height
  xEnd = cDibTo.Width - 1
  For y = 0 To cDibTo.Height - 1
    fY = y * yScale
    ifY = Int(fY)
    dy = fY - ifY
    For x = 0 To xEnd
      FX = x * xScale
      ifX = Int(FX)
      dX = FX - ifX
      ifX = ifX * 3
      ' Interpolate using the four nearest pixels in the source
      B1 = bDibFrom(ifX, ifY): G1 = bDibFrom(ifX + 1, ifY): R1 = bDibFrom(ifX + 2, ifY)
      B2 = bDibFrom(ifX + 3, ifY): G2 = bDibFrom(ifX + 4, ifY): R2 = bDibFrom(ifX + 5, ifY)
      B3 = bDibFrom(ifX, ifY + 1): G3 = bDibFrom(ifX + 1, ifY + 1): R3 = bDibFrom(ifX + 2, ifY + 1)
      B4 = bDibFrom(ifX + 3, ifY + 1): G4 = bDibFrom(ifX + 4, ifY + 1): R4 = bDibFrom(ifX + 5, ifY + 1)
      ' Interpolate in x direction:
      iR1 = R1 * (1 - dy) + R3 * dy: iG1 = G1 * (1 - dy) + G3 * dy: iB1 = B1 * (1 - dy) + B3 * dy
      iR2 = R2 * (1 - dy) + R4 * dy: iG2 = G2 * (1 - dy) + G4 * dy: iB2 = B2 * (1 - dy) + B4 * dy
      ' Interpolate in y:
      R = iR1 * (1 - dX) + iR2 * dX: G = iG1 * (1 - dX) + iG2 * dX: b = iB1 * (1 - dX) + iB2 * dX
      ' Set output:
      If (R < 0) Then R = 0
      If (R > 255) Then R = 255
      If (G < 0) Then G = 0
      If (G > 255) Then G = 255
      If (b < 0) Then b = 0
      If (b > 255) Then b = 255
      xOut = x * 3
      bDibTo(xOut, y) = b
      bDibTo(xOut + 1, y) = G
      bDibTo(xOut + 2, y) = R
    Next x
  Next y

  'Clear the temporary array descriptor
  'This is necessary under NT4.
  CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
  CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4
  
  Create lNewWidth, lNewHeight
  DC2Object cDibTo.hDC
  m_TC = GetPixel(m_hDC, 0, 0)
  
  Set cDibTo = Nothing
End Sub

Public Sub FX_AddImage(ByRef cFrom As clsDIBSection, ByVal destX As Long, ByVal destY As Long, ByVal Transparent As FX_AddImage_Transparency)
  'Add a source DIB to the current DIB at zero-based (destX,destY)
  'If "Transparent" then add it transparently using the "TransparentColor" property of the source DIB
  Dim picTo()   As Byte
  Dim PicFrom() As Byte
  Dim tSA       As SAFEARRAY2D
  Dim tSA2      As SAFEARRAY2D
  Dim x         As Long
  Dim y         As Long
  Dim yMax      As Long
  Dim xMax      As Long
  Dim X2        As Long
  Dim Y2        As Long
  Dim xStart    As Long
  Dim rR        As Long
  Dim rG        As Long
  Dim rB        As Long
  Dim tR        As Long
  Dim tG        As Long
  Dim tB        As Long
  Dim lGray1    As Long
  Dim lGray2    As Long
  
  'Have the local matrix point to bitmap pixels
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = Me.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = Me.BytesPerScanLine
    .pvData = Me.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(picTo), VarPtr(tSA), 4
  'Have the local matrix point to bitmap pixels
  With tSA2
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = cFrom.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = cFrom.BytesPerScanLine
    .pvData = cFrom.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(PicFrom), VarPtr(tSA2), 4
  'Only usefull if "Transparent":
  tR = cFrom.TransparentColor And &HFF&
  tG = (cFrom.TransparentColor And &HFF00&) \ &H100&
  tB = (cFrom.TransparentColor And &HFF0000) \ &H10000
  If (Transparent = FXAT_AddAlphaBlended) Then
    'transparent color is minus 6%
    tR = tR - (tR * 0.1)
    tG = tG - (tG * 0.1)
    tB = tB - (tB * 0.1)
    lGray1 = tB + tG * &H100 + tR * &H10000
  End If
  'Do the addition:
  xStart = destX * 3
  yMax = cFrom.Height - 1
  xMax = (cFrom.Width - 1) * 3
  For x = 0 To xMax Step 3
    For y = 0 To yMax
      If (Transparent = FXAT_AddDarkest) Then
        X2 = xStart + x: Y2 = destY + y
        lGray1 = (222& * picTo(X2 + 1, Y2) + 707& * picTo(X2 + 1, Y2) + 71& * picTo(X2, Y2))
        lGray2 = (222& * PicFrom(x + 1, y) + 707& * PicFrom(x + 1, y) + 71& * PicFrom(x, y))
        If (lGray1 > lGray2) Then
          picTo(X2, Y2) = PicFrom(x, y)
          picTo(X2 + 1, Y2) = PicFrom(x + 1, y)
          picTo(X2 + 2, Y2) = PicFrom(x + 2, y)
        End If
      ElseIf (Transparent = FXAT_AddTransparent) Then
        'Transparent add
        If (PicFrom(x + 2, y) <> tR) Or (PicFrom(x + 1, y) <> tG) Or (PicFrom(x, y) <> tB) Then
          X2 = xStart + x: Y2 = destY + y
          picTo(X2, Y2) = PicFrom(x, y)
          picTo(X2 + 1, Y2) = PicFrom(x + 1, y)
          picTo(X2 + 2, Y2) = PicFrom(x + 2, y)
        End If
      ElseIf (Transparent = FXAT_Combine) Then
        'Combine add:
        X2 = xStart + x: Y2 = destY + y
        rR = PicFrom(x, y) / 2
        rG = PicFrom(x + 1, y) / 2
        rB = PicFrom(x + 2, y) / 2
        picTo(X2, Y2) = (rB + picTo(X2, Y2)) \ 2
        picTo(X2 + 1, Y2) = (rG + picTo(X2 + 1, Y2)) \ 2
        picTo(X2 + 2, Y2) = (rR + picTo(X2 + 2, Y2)) \ 2
      ElseIf (Transparent = FXAT_CombineTransparent) Then
        'Combine add with transparency:
        If (PicFrom(x + 2, y) <> tR) Or (PicFrom(x + 1, y) <> tG) Or (PicFrom(x, y) <> tB) Then
          X2 = xStart + x: Y2 = destY + y
          rR = PicFrom(x, y) '/ 2
          rG = PicFrom(x + 1, y) ' / 2
          rB = PicFrom(x + 2, y) '/ 2
          picTo(X2, Y2) = (rB + picTo(X2, Y2)) \ 2
          picTo(X2 + 1, Y2) = (rG + picTo(X2 + 1, Y2)) \ 2
          picTo(X2 + 2, Y2) = (rR + picTo(X2 + 2, Y2)) \ 2
        End If
      ElseIf (Transparent = FXAT_AddAlphaBlended) Then
        '50% blended
        lGray2 = CLng(PicFrom(x, y)) * &H10000 + CLng(PicFrom(x + 1, y)) * &H100 + PicFrom(x + 2, y)
        If (lGray2 < lGray1) Then
          X2 = xStart + x: Y2 = destY + y
          rB = (CLng(picTo(X2, Y2)) + CLng(PicFrom(x, y))) / 2
          rG = (CLng(picTo(X2 + 1, Y2)) + CLng(PicFrom(x + 1, y))) / 2
          rR = (CLng(picTo(X2 + 2, Y2)) + CLng(PicFrom(x + 2, y))) / 2
          picTo(X2, Y2) = rB
          picTo(X2 + 1, Y2) = rG
          picTo(X2 + 2, Y2) = rR
        End If
      Else
        'Normal add:
        X2 = xStart + x: Y2 = destY + y
        picTo(X2, Y2) = PicFrom(x, y)
        picTo(X2 + 1, Y2) = PicFrom(x + 1, y)
        picTo(X2 + 2, Y2) = PicFrom(x + 2, y)
      End If
    Next y
  Next x
  'Clear the temporary array descriptor without destroying the local temporary array
  CopyMemory ByVal VarPtrArray(picTo), 0&, 4
  CopyMemory ByVal VarPtrArray(PicFrom), 0&, 4
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub
  
Public Sub FX_GrayScale()
  'Gray scale the current DIB object using standard intensity components.
  'see http://www.dcs.ed.ac.uk/~mxr/gfx/faqs/colourspace.faq for details.
  Dim bDib()    As Byte
  Dim x         As Long
  Dim y         As Long
  Dim xMax      As Long
  Dim yMax      As Long
  Dim lB        As Long
  Dim lG        As Long
  Dim lR        As Long
  Dim lGray     As Long
  Dim tSA       As SAFEARRAY2D
        
  'Have the local matrix point to bitmap pixels
  With tSA
   .cbElements = 1
   .cDims = 2
   .Bounds(0).lLbound = 0
   .Bounds(0).cElements = Me.Height
   .Bounds(1).lLbound = 0
   .Bounds(1).cElements = Me.BytesPerScanLine
   .pvData = Me.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
  
  'Do the grayscale:
  yMax = Me.Height - 1
  xMax = (Me.Width - 1) * 3
  For x = 0 To xMax Step 3
    For y = 0 To yMax
      lB = bDib(x, y)
      lG = bDib(x + 1, y)
      lR = bDib(x + 2, y)
      'But now all people *should* use the most accurate, it means ITU standard:
      lGray = (222 * lR + 707 * lG + 71 * lB) / 1000
      bDib(x, y) = lGray
      bDib(x + 1, y) = lGray
      bDib(x + 2, y) = lGray
    Next y
  Next x
    
  'Clear the temporary array descriptor without destroying the local temporary array
  CopyMemory ByVal VarPtrArray(bDib), 0&, 4
  
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Public Sub FX_Darken(ByVal lAmount As Long)
  'Fade the current DIB object to black
  Dim bDib()    As Byte
  Dim x         As Long
  Dim y         As Long
  Dim xMax      As Long
  Dim yMax      As Long
  Dim lR        As Long
  Dim lG        As Long
  Dim lB        As Long
  Dim tSA       As SAFEARRAY2D
  
  'Have the local matrix point to bitmap pixels
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = Me.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = Me.BytesPerScanLine
    .pvData = Me.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
  
  'Do the fade:
  yMax = Me.Height - 1
  xMax = (Me.Width - 1) * 3
  For x = 0 To xMax Step 3
    For y = 0 To yMax
      lB = lAmount * bDib(x, y) \ 255
      lG = lAmount * bDib(x + 1, y) \ 255
      lR = lAmount * bDib(x + 2, y) \ 255
      bDib(x, y) = lB
      bDib(x + 1, y) = lG
      bDib(x + 2, y) = lR
    Next y
  Next x
    
  'Clear the temporary array descriptor without destroying the local temporary array
  CopyMemory ByVal VarPtrArray(bDib), 0&, 4
  
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Public Sub FX_Lighten(ByVal lAmount As Long)
  'Lighten the current DIB object
  Dim bDib()    As Byte
  Dim x         As Long
  Dim y         As Long
  Dim xMax      As Long
  Dim yMax      As Long
  Dim lR        As Long
  Dim lG        As Long
  Dim lB        As Long
  Dim tSA       As SAFEARRAY2D
  Dim H         As Single
  Dim S         As Single
  Dim L         As Single
    
  'Have the local matrix point to bitmap pixels
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = Me.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = Me.BytesPerScanLine
    .pvData = Me.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
  
  'Do the fade:
  yMax = Me.Height - 1
  xMax = (Me.Width - 1) * 3
  For x = 0 To xMax Step 3
    For y = 0 To yMax
      RGBToHLS bDib(x + 2, y), bDib(x + 1, y), bDib(x, y), H, S, L
      L = L * (1 + (lAmount / 100))
      If (L > 1) Then L = 1
      HLSToRGB H, S, L, lR, lG, lB
      bDib(x, y) = lB
      bDib(x + 1, y) = lG
      bDib(x + 2, y) = lR
    Next y
  Next x
    
  'Clear the temporary array descriptor without destroying the local temporary array
  CopyMemory ByVal VarPtrArray(bDib), 0&, 4
  
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Public Sub FX_DropShadow(ByRef cResult As clsDIBSection, ByVal ShadowDepth As FX_DropShadow_ShadowDepth)
  Dim bDibFrom() As Byte
  Dim bDibTo()   As Byte
  Dim tSA        As SAFEARRAY2D
  Dim tSA2       As SAFEARRAY2D
  Dim x          As Long
  Dim y          As Long
  Dim rgbOffset  As Long
  Dim xOffset    As Long
  Dim R          As Long
  Dim G          As Long
  Dim b          As Long
  Dim i          As Long
  Dim j          As Long
  Dim yMax       As Long
  Dim xMax       As Long
  Dim rR         As Long
  Dim rB         As Long
  Dim rG         As Long
  Dim Offset     As Long
  Dim Weight     As Long
  Dim Filter()   As Long
       
  'have the local matrix point to bitmap pixels
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = Me.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = Me.BytesPerScanLine
    .pvData = Me.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibFrom), VarPtr(tSA), 4
        
  'have the buffer matrix point to bitmap pixels
  With tSA2
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = cResult.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = cResult.BytesPerScanLine
    .pvData = cResult.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibTo), VarPtr(tSA2), 4
  
  'Build filter array:
  'Example: ShadowDepth = 3 => 3\2=1 => 1 pixel around middle pixel
  'o o o
  'o x o
  'o o o
  Weight = 0: Offset = ShadowDepth \ 2 'Odd number between 1..9
  ReDim Filter(-Offset To Offset, -Offset To Offset)
  For i = -Offset To Offset
    For j = -Offset To Offset
      Filter(i, j) = 1
      Weight = Weight + Filter(i, j)
    Next j
  Next i
    
  'Perform blur filter on bDibFrom into bDibTo
  rgbOffset = Offset * 3
  yMax = Me.Height - 1 - Offset
  xMax = (Me.Width - 1) * 3 - rgbOffset
      
  For x = rgbOffset To xMax Step 3
    For y = Offset To yMax
      R = 0: G = 0: b = 0
      For i = -Offset To Offset
        xOffset = i * 3
        For j = -Offset To Offset
          R = R + Filter(i, j) * bDibFrom(x + 2 + xOffset, y + j)
          G = G + Filter(i, j) * bDibFrom(x + 1 + xOffset, y + j)
          b = b + Filter(i, j) * bDibFrom(x + 0 + xOffset, y + j)
        Next j
      Next i
      rR = R \ Weight: rG = G \ Weight: rB = b \ Weight
      If (rR < 0) Then rR = 0
      If (rG < 0) Then rG = 0
      If (rB < 0) Then rB = 0
      If (rR > 255) Then rR = 255
      If (rG > 255) Then rG = 255
      If (rB > 255) Then rB = 255
      bDibTo(x + 2, y) = rR: bDibTo(x + 1, y) = rG: bDibTo(x, y) = rB
    Next y
  Next x
        
  'Clear the temporary array descriptors without destroying the local temporary arrays
  CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
  CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4
  
  Erase Filter
          
  m_TC = GetPixel(m_hDC, 0, 0)
End Sub

Private Function CheckVal(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long) As Long
  'lValue has to be between lMin..lMax
  If (lValue < lMin) Then
    CheckVal = lMin
  ElseIf (lValue > lMax) Then
    CheckVal = lMax
  Else
    CheckVal = lValue
  End If
End Function

Private Sub RGBToHLS(ByVal R As Long, ByVal G As Long, ByVal b As Long, H As Single, S As Single, L As Single)
  Dim Max   As Single
  Dim Min   As Single
  Dim delta As Single
  Dim rR    As Single
  Dim rG    As Single
  Dim rB    As Single

  rR = R / 255: rG = G / 255: rB = b / 255

  '{Given: rgb each in [0,1].
  ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
  Max = Maximum(rR, rG, rB)
  Min = Minimum(rR, rG, rB)
  L = (Max + Min) / 2    '{This is the lightness}
  '{Next calculate saturation}
  If (Max = Min) Then
    'Achromatic case
    S = 0
    H = 0
  Else
    'Chromatic case
    'First calculate the saturation.
    If (L <= 0.5) Then
      S = (Max - Min) / (Max + Min)
    Else
      S = (Max - Min) / (2 - Max - Min)
    End If
    'Next calculate the hue.
    delta = Max - Min
    If (rR = Max) Then
      H = (rG - rB) / delta     'Resulting color is between yellow and magenta
    ElseIf rG = Max Then
      H = 2 + (rB - rR) / delta 'Resulting color is between cyan and yellow
    ElseIf rB = Max Then
      H = 4 + (rR - rG) / delta 'Resulting color is between magenta and cyan
    End If
  End If
End Sub

Private Sub HLSToRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single, R As Long, G As Long, b As Long)
  Dim rR  As Single
  Dim rG  As Single
  Dim rB  As Single
  Dim Min As Single
  Dim Max As Single

  If (S = 0) Then
    'Achromatic case:
    rR = L: rG = L: rB = L
  Else
    'Chromatic case:
    ' delta = Max-Min
    If (L <= 0.5) Then
      's = (Max - Min) / (Max + Min)
      ' Get Min value:
      Min = L * (1 - S)
    Else
      's = (Max - Min) / (2 - Max - Min)
      ' Get Min value:
      Min = L - S * (1 - L)
    End If
    ' Get the Max value:
    Max = 2 * L - Min
    ' Now depending on sector we can evaluate the h,l,s:
    If (H < 1) Then
      rR = Max
      If (H < 0) Then
        rG = Min
        rB = rG - H * (Max - Min)
      Else
        rB = Min
        rG = H * (Max - Min) + rB
      End If
    ElseIf (H < 3) Then
      rG = Max
      If (H < 2) Then
        rB = Min
        rR = rB - (H - 2) * (Max - Min)
      Else
        rR = Min
        rB = (H - 2) * (Max - Min) + rR
      End If
    Else
      rB = Max
      If (H < 4) Then
        rR = Min
        rG = rR - (H - 4) * (Max - Min)
      Else
        rG = Min
        rR = (H - 4) * (Max - Min) + rG
      End If
    End If
  End If
  R = rR * 255
  G = rG * 255
  b = rB * 255
End Sub

Private Function Maximum(Red As Single, Green As Single, Blue As Single) As Single
  If (Red > Green) Then
    If (Red > Blue) Then Maximum = Red Else Maximum = Blue
  Else
    If (Blue > Green) Then Maximum = Blue Else Maximum = Green
  End If
End Function

Private Function Minimum(Red As Single, Green As Single, Blue As Single) As Single
  If (Red < Green) Then
    If (Red < Blue) Then Minimum = Red Else Minimum = Blue
  Else
    If (Blue < Green) Then Minimum = Blue Else Minimum = Green
  End If
End Function
