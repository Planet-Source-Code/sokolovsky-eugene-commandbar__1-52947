VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCommandBarImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" ( _
        ByVal hIml As Long, ByVal i As Long, _
        ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840
Private Const ILC_COLOR = &H0
Private Const ILC_COLOR32 = &H20
Private Const ILC_MASK = &H1&

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long

Private Type ICONINFO
   fIcon As Long
   xHotspot As Long
   yHotspot As Long
   hBmMask As Long
   hbmColor As Long
End Type
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Device dependent Bitmap structure:
Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_FIXED = &H0
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "msimg32.dll" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hDcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Enum EIconProcessorStyle
   eIconStandard = 0
   eIconDIsabled = 1
   eIconNonHighlighted = 2
   eIconHighlighted = 3
End Enum

Private m_ptrVB6ImageList As Long
Private m_hIml As Long
Private m_lIconWidth As Long
Private m_lIconHeight As Long

Private m_lChunkSize As Long
Private m_lCurrentChunks As Long

Private m_hDC As Long
Private m_hDib As Long
Private m_lPtr As Long
Private m_hBmpOld As Long
Private m_tBI As BITMAPINFO
Private m_lWidth As Long
Private m_lHeight As Long

Private Type tIconInfo
   iIconIndex As Long
   xPosition As Long
   eStyle As EIconProcessorStyle
End Type
Private m_tIcon() As tIconInfo
Private m_iIconCount As Long

Private m_oDisabledColor As OLE_COLOR
Private m_oHighlightColor As OLE_COLOR
Private m_lAlpha As Long

Friend Property Get hIml() As Long
   hIml = m_hIml
End Property
Friend Property Get VB6ImageListPtr() As Long
   VB6ImageListPtr = m_ptrVB6ImageList
End Property
Public Property Get IconWidth() As Long
   IconWidth = m_lIconWidth
End Property
Public Property Get IconHeight() As Long
   IconHeight = m_lIconHeight
End Property

Public Sub DrawStrip(ByVal lhDCTO As Long, ByVal x As Long, ByVal y As Long)
    BitBlt lhDCTO, x, y, m_lWidth, m_lHeight, m_hDC, 0, 0, vbSrcCopy
End Sub

Public Property Get DisabledColor() As OLE_COLOR
   DisabledColor = m_oDisabledColor
End Property
Public Property Let DisabledColor(ByVal oColor As OLE_COLOR)
   If Not (m_oDisabledColor = oColor) Then
      m_oDisabledColor = oColor
      If (m_iIconCount > 0) Then
         Dim i As Long
         For i = 1 To m_iIconCount
            If (m_tIcon(i).eStyle = eIconDIsabled) Then
               CreateProcessedIcon m_tIcon(i).iIconIndex, eIconDIsabled
            End If
         Next i
      End If
   End If
End Property

Public Property Get HighlightColor() As OLE_COLOR
   HighlightColor = m_oHighlightColor
End Property
Public Property Let HighlightColor(ByVal oColor As OLE_COLOR)
   If Not (m_oHighlightColor = oColor) Then
      m_oHighlightColor = oColor
      If (m_iIconCount > 0) Then
         Dim i As Long
         For i = 1 To m_iIconCount
            If (m_tIcon(i).eStyle = eIconHighlighted) Then
               CreateProcessedIcon m_tIcon(i).iIconIndex, eIconHighlighted
            End If
         Next i
      End If
   End If
End Property

Public Property Get Count() As Long
   Count = m_iIconCount
End Property

Public Property Get IconIndexOf( _
      ByVal lIndex As Long _
   ) As Long
   IconIndexOf = m_tIcon(lIndex).iIconIndex
End Property

Public Property Get IndexOf( _
      ByVal lIconIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   ) As Long
Dim i As Long
Dim lIndex As Long

   For i = 1 To m_iIconCount
      If (m_tIcon(i).iIconIndex = lIconIndex) Then
         If (m_tIcon(i).eStyle = eStyle) Then
            IndexOf = i
            Exit For
         End If
      End If
   Next i
      
End Property

Private Sub CreateProcessedIcon( _
      ByVal lIconIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim lIndex As Long

   lIndex = IndexOf(lIconIndex, eStyle)
   If (lIndex = 0) Then
   
      If (m_iIconCount + 2) > m_lCurrentChunks Then
         ChunkResize
      End If
      
      m_iIconCount = m_iIconCount + 1
      lIndex = m_iIconCount
      
      ReDim Preserve m_tIcon(1 To m_iIconCount) As tIconInfo
      m_tIcon(lIndex).iIconIndex = lIconIndex
      m_tIcon(lIndex).eStyle = eStyle
      m_tIcon(lIndex).xPosition = (lIndex - 1) * m_lIconWidth
      
      If Not (m_ptrVB6ImageList = 0) Then
         CreateProcessedIconFromRubbishImageList lIndex, eStyle
      Else
         CreateProcessedIconFromRealImageList lIndex, eStyle
      End If
      
   End If
   
   
End Sub

Public Sub Draw( _
      ByVal lHDC As Long, _
      ByVal lIconIndex As Long, _
      ByVal eStyle As EIconProcessorStyle, _
      ByVal lX As Long, _
      ByVal lY As Long, _
      Optional ByVal lWidth As Long = -1, _
      Optional ByVal lHeight As Long = -1 _
   )
   
   If (lWidth < 0) Then lWidth = m_lIconWidth
   If (lHeight < 0) Then lHeight = m_lIconHeight
   If (lWidth > m_lIconWidth) Then lWidth = m_lIconWidth
   If (lHeight > m_lIconHeight) Then lHeight = m_lIconHeight
   
   If (eStyle = eIconStandard) Then
      
      If Not (m_ptrVB6ImageList = 0) Then
         Dim o As Object
         On Error Resume Next
         Set o = ObjectFromPtr(m_ptrVB6ImageList)
         If Not (o Is Nothing) Then
            Dim scaleIconX As Single
            Dim scaleIconY As Single
            scaleIconX = o.Parent.ScaleX(lX, vbPixels, o.Parent.ScaleMode)
            scaleIconY = o.Parent.ScaleY(lY, vbPixels, o.Parent.ScaleMode)
             o.ListImages(lIconIndex + 1).Draw lHDC, scaleIconX, scaleIconY, ILD_TRANSPARENT
         End If
         On Error GoTo 0
      Else
         ImageList_Draw _
             m_hIml, _
             lIconIndex, _
             lHDC, _
             lX, _
             lY, _
             ILD_TRANSPARENT
      End If
   
   Else
      
      CreateProcessedIcon lIconIndex, eStyle
   
      Dim lBlend As Long
      Dim bf As BLENDFUNCTION
      bf.BlendOp = AC_SRC_OVER
      bf.BlendFlags = 0
      bf.SourceConstantAlpha = 255
      bf.AlphaFormat = AC_SRC_ALPHA
      CopyMemory lBlend, bf, 4
      
      Dim lR As Long
      Dim srcX As Long
      lR = IndexOf(lIconIndex, eStyle)
      If (lR > 0) Then
         srcX = m_tIcon(lR).xPosition
         If (goodSystem) Then
            lR = AlphaBlend( _
               lHDC, _
               lX, lY, lWidth, lHeight, _
               m_hDC, _
               srcX, 0, lWidth, lHeight, _
               lBlend)
            If (lR = 0) Then
               Debug.Print "An error drawing image...", m_lWidth, m_lHeight
            End If
         Else
            BitBlt m_hDC, (m_iIconCount * m_lIconWidth), 0, lWidth, lHeight, lHDC, lX, lY, vbSrcCopy
            CodeAlphaBlend srcX, lWidth, lHeight
            BitBlt lHDC, lX, lY, lWidth, lHeight, m_hDC, (m_iIconCount * m_lIconWidth), 0, vbSrcCopy
         End If
      End If
      
   End If
End Sub

Private Property Get goodSystem() As Boolean
   goodSystem = Is2000OrAbove
End Property

Private Sub CodeAlphaBlend( _
      ByVal srcX As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
Dim tSA As SAFEARRAY2D
Dim bDib() As Byte
Dim xSrc As Long
Dim xDst As Long
Dim xStartSrc As Long
Dim xEndSrc As Long
Dim xStartDst As Long
Dim xEndDst As Long
Dim y As Long
Dim lAlpha As Long
   
   With tSA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lWidth * 4
      .pvData = m_lPtr
   End With
   CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

   xStartSrc = srcX * 4
   xEndSrc = xStartSrc + lWidth * 4 - 4
   xDst = m_iIconCount * m_lIconWidth * 4
   
   For xSrc = xStartSrc To xEndSrc Step 4
      For y = 0 To m_lHeight - 1
         If (bDib(xSrc + 3, y) = 0) Then
            ' transparent
         Else
            ' 'alpha blend'
            lAlpha = bDib(xSrc + 3, y)
            bDib(xDst, y) = bDib(xSrc, y) + bDib(xDst, y) * (255 - lAlpha) / 255&
            bDib(xDst + 1, y) = bDib(xSrc + 1, y) + bDib(xDst + 1, y) * (255 - lAlpha) / 255&
            bDib(xDst + 2, y) = bDib(xSrc + 2, y) + bDib(xDst + 2, y) * (255 - lAlpha) / 255&
         End If
      Next y
      xDst = xDst + 4
   Next xSrc
   
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4

End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property


Private Sub CreateProcessedIconFromRubbishImageList( _
      ByVal lIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim ilsIcons As Object
Dim lHDC As Long
Dim hDib As Long
Dim hBmpOld As Long
Dim lPtr As Long
Dim tBI As BITMAPINFO
Dim lhDCComp As Long
Dim lhWndD As Long
Dim tR As RECT
Dim hBr As Long

   lhWndD = GetDesktopWindow()
   lhDCComp = GetDC(lhWndD)

   With tBI.bmiHeader
      .biSize = Len(tBI.bmiHeader)
      .biWidth = m_lIconWidth
      .biHeight = m_lIconHeight
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = BI_RGB
      .biSizeImage = .biWidth * .biHeight * 4 ' 32 bpp
   End With
   hDib = CreateDIBSection( _
           lhDCComp, _
           tBI, _
           DIB_RGB_COLORS, _
           lPtr, _
           0, 0)
           
   If Not (hDib = 0) Then
      lHDC = CreateCompatibleDC(lhDCComp)
      hBmpOld = SelectObject(lHDC, hDib)
      hBr = CreateSolidBrush(&H10201)
      tR.right = m_lIconWidth
      tR.bottom = m_lIconHeight
      FillRect lHDC, tR, hBr
      DeleteObject hBr
   
      On Error GoTo ErrorHandler
      Set ilsIcons = ObjectFromPtr(m_ptrVB6ImageList)
      ilsIcons.ListImages(m_tIcon(lIndex).iIconIndex + 1).Draw lHDC, 0, 0, 1
      
      SelectObject lHDC, hBmpOld
      DeleteObject lHDC
      
      CreateFromDIB lPtr, lIndex, eStyle
      
      DeleteObject hDib
      
      ReleaseDC lhWndD, lhDCComp
   End If
      
   Exit Sub
ErrorHandler:
   Exit Sub
End Sub

Private Sub CreateFromDIB( _
      ByVal lPtr As Long, _
      ByVal lIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim tSAColor As SAFEARRAY2D
Dim bDibColor() As Byte
Dim tSADest As SAFEARRAY2D
Dim bDibDest() As Byte
Dim x As Long
Dim xEnd As Long
Dim y As Long
Dim xDst As Long
Dim xDstInit As Long
Dim lAlpha As Long
Dim lR As Long
Dim lG As Long
Dim lB As Long
Dim lGrey As Long
Dim lResR As Long
Dim lResB As Long
Dim lResG As Long
Dim lDisabledColor As Long
Dim lHighR As Long
Dim lHighB As Long
Dim lHighG As Long
Dim lHighlightColor As Long
Dim bAllWhite As Boolean
   
   OleTranslateColor m_oDisabledColor, 0, lDisabledColor
   lResR = lDisabledColor And &HFF&
   lResG = (lDisabledColor And &HFF00&) \ &H100&
   lResB = (lDisabledColor And &HFF0000) \ &H10000
   
   OleTranslateColor m_oHighlightColor, 0, lHighlightColor
   lHighR = lHighlightColor And &HFF&
   lHighG = (lHighlightColor And &HFF00&) \ &H100&
   lHighB = (lHighlightColor And &HFF0000) \ &H10000
   
   With tSAColor
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lIconHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lIconWidth * 4
      .pvData = lPtr
   End With
   CopyMemory ByVal VarPtrArray(bDibColor()), VarPtr(tSAColor), 4
      
   With tSADest
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lWidth * 4
      .pvData = m_lPtr
   End With
   CopyMemory ByVal VarPtrArray(bDibDest()), VarPtr(tSADest), 4

   xDst = m_tIcon(lIndex).xPosition * 4
   xDstInit = xDst
   xEnd = (m_lIconWidth - 1) * 4
   
   bAllWhite = True
   For y = 0 To m_lIconHeight - 1
      For x = 0 To xEnd Step 4
         If (bDibColor(x, y) = &H1 And bDibColor(x + 1, y) = &H2 And bDibColor(x + 2, y) = &H1) Then
         ElseIf ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) _
            Or ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
         ElseIf (bDibColor(x, y) = 255 And bDibColor(x + 1, y) = 255 And bDibColor(x + 2, y) = 255) Then
         Else
            bAllWhite = False
         End If
      Next x
   Next y
   
   
   For y = 0 To m_lIconHeight - 1
      For x = 0 To xEnd Step 4
      
         If (bDibColor(x, y) = &H1 And bDibColor(x + 1, y) = &H2 And bDibColor(x + 2, y) = &H1) Then
            ' output is transparent
            lR = 0
            lG = 0
            lB = 0
            lAlpha = 0
         Else
            lAlpha = 255
            
            Select Case eStyle
            Case eIconHighlighted
               If ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) Then
                  lR = lHighR
                  lG = lHighG
                  lB = lHighB
                  lAlpha = lAlpha * 90& \ 255&
               ElseIf ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
                  lR = lHighR
                  lG = lHighG
                  lB = lHighB
                  lAlpha = lAlpha * 90& \ 255&
               Else
                  If (bAllWhite) Then
                     lAlpha = 255
                  End If
                  lR = bDibColor(x + 2, y)
                  lG = bDibColor(x + 1, y)
                  lB = bDibColor(x, y)
               End If
               
            Case eIconNonHighlighted
               lR = bDibColor(x + 2, y)
               lG = bDibColor(x + 1, y)
               lB = bDibColor(x, y)
               If (bAllWhite) Then
                  lAlpha = 255
               End If
               If ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) _
                  Or ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
                  lAlpha = lAlpha * 50& / 255&
               End If
            
            Case eIconDIsabled
               lR = lResR
               lG = lResG
               lB = lResB
               If (bAllWhite) And (bDibColor(x + 2, y) = 255) And (bDibColor(x + 1, y) = 255) And (bDibColor(x, y) = 255) Then
                  lAlpha = 255
               Else
                  ' Decrease alpha in proportion to the grey value:
                  lGrey = (222& * bDibColor(x + 2, y) + 707& * bDibColor(x + 1, y) + 71& * bDibColor(x, y)) / 1000&
                  lAlpha = lAlpha * (255& - lGrey) / 255&
               End If
               
            End Select
            
            ' Premultiply alpha:
            lR = lR * lAlpha \ 255
            lG = lG * lAlpha \ 255
            lB = lB * lAlpha \ 255
                     
         End If
         
         bDibDest(xDst, y) = lB
         bDibDest(xDst + 1, y) = lG
         bDibDest(xDst + 2, y) = lR
         bDibDest(xDst + 3, y) = lAlpha
         
         xDst = xDst + 4
      Next x
      
      xDst = xDstInit
      
   Next y

   CopyMemory ByVal VarPtrArray(bDibDest), 0&, 4
   CopyMemory ByVal VarPtrArray(bDibColor), 0&, 4

   

End Sub

Private Sub CreateProcessedIconFromRealImageList( _
      ByVal lIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim hIcon As Long
   hIcon = ImageList_GetIcon(m_hIml, m_tIcon(lIndex).iIconIndex, 0)
   If Not (hIcon = 0) Then
      CreateFromhIcon hIcon, lIndex, eStyle
      DestroyIcon hIcon
   End If
End Sub

Private Sub CreateFromhIcon( _
      ByVal hIcon As Long, _
      ByVal lIndex As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim tII As ICONINFO
Dim lHDC As Long
Dim lhWndD As Long
Dim lR As Long

   lhWndD = GetDesktopWindow()
   lHDC = GetDC(lhWndD)

   GetIconInfo hIcon, tII
   If Not (tII.hbmColor = 0) And Not (tII.hBmMask = 0) Then
   
      ' Get the colour bitmap as a 32bpp DIB:
      Dim tBMColor As BITMAP
      Dim lSize As Long
      GetObjectAPI tII.hbmColor, Len(tBMColor), tBMColor
      lSize = tBMColor.bmWidth * tBMColor.bmHeight * 4
      Dim hMemColor As Long
      Dim lPtrColor As Long
      hMemColor = LocalAlloc(GPTR, lSize)
      If Not (hMemColor = 0) Then
         lPtrColor = LocalLock(hMemColor)
         If Not (lPtrColor = 0) Then
            Dim tBIColor As BITMAPINFO
            With tBIColor.bmiHeader
               .biSize = Len(tBIColor.bmiHeader)
               .biWidth = tBMColor.bmWidth
               .biHeight = tBMColor.bmHeight
               .biPlanes = 1
               .biBitCount = 32
               .biCompression = BI_RGB
               .biSizeImage = lSize
            End With
            lR = GetDIBits(lHDC, tII.hbmColor, 0, tBMColor.bmHeight, ByVal lPtrColor, tBIColor, DIB_RGB_COLORS)
      
            ' Get the mask bitmap as a 32bpp DIB:
            Dim tBMMask As BITMAP
            GetObjectAPI tII.hBmMask, Len(tBMMask), tBMMask
            Dim hMemMask As Long
            Dim lPtrMask As Long
            hMemMask = LocalAlloc(GPTR, lSize)
            If Not (hMemMask = 0) Then
               lPtrMask = LocalLock(hMemMask)
               If Not (lPtrMask = 0) Then
                  Dim tBIMask As BITMAPINFO
                  With tBIMask.bmiHeader
                     .biSize = Len(tBIMask.bmiHeader)
                     .biWidth = tBMMask.bmWidth
                     .biHeight = tBMMask.bmHeight
                     .biPlanes = 1
                     .biBitCount = 32
                     .biCompression = BI_RGB
                     .biSizeImage = lSize
                  End With
                  lR = GetDIBits(lHDC, tII.hBmMask, 0, tBMMask.bmHeight, ByVal lPtrMask, tBIMask, DIB_RGB_COLORS)
            
                  CreateFromIconDIBs lIndex, lPtrColor, lPtrMask, eStyle
               
                  LocalUnlock hMemMask
               End If
               LocalFree hMemMask
            End If
            
            LocalUnlock hMemColor

         End If
         LocalFree hMemColor

      End If
      
      
      
   End If
   
   If Not (tII.hbmColor = 0) Then
      DeleteObject tII.hbmColor
   End If
   If Not (tII.hBmMask = 0) Then
      DeleteObject tII.hBmMask
   End If
   ReleaseDC lhWndD, lHDC
   
End Sub

Private Sub CreateFromIconDIBs( _
      ByVal lIndex As Long, _
      ByVal lPtrColor As Long, _
      ByVal lPtrMask As Long, _
      ByVal eStyle As EIconProcessorStyle _
   )
Dim tSAColor As SAFEARRAY2D
Dim bDibColor() As Byte
Dim tSAMask As SAFEARRAY2D
Dim bDibMask() As Byte
Dim tSADest As SAFEARRAY2D
Dim bDibDest() As Byte
Dim x As Long
Dim xEnd As Long
Dim y As Long
Dim xDst As Long
Dim xDstInit As Long
Dim lAlpha As Long
Dim lR As Long
Dim lG As Long
Dim lB As Long
Dim bAllZero As Boolean
Dim bAllWhite As Boolean
Dim lGrey As Long
Dim lResR As Long
Dim lResB As Long
Dim lResG As Long
Dim lDisabledColor As Long
Dim lHighR As Long
Dim lHighB As Long
Dim lHighG As Long
Dim lHighlightColor As Long
   
   OleTranslateColor m_oDisabledColor, 0, lDisabledColor
   lResR = lDisabledColor And &HFF&
   lResG = (lDisabledColor And &HFF00&) \ &H100&
   lResB = (lDisabledColor And &HFF0000) \ &H10000
   
   OleTranslateColor m_oHighlightColor, 0, lHighlightColor
   lHighR = lHighlightColor And &HFF&
   lHighG = (lHighlightColor And &HFF00&) \ &H100&
   lHighB = (lHighlightColor And &HFF0000) \ &H10000
   
   With tSAColor
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lIconHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lIconWidth * 4
      .pvData = lPtrColor
   End With
   CopyMemory ByVal VarPtrArray(bDibColor()), VarPtr(tSAColor), 4
   
   With tSAMask
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lIconHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lIconWidth * 4
      .pvData = lPtrMask
   End With
   CopyMemory ByVal VarPtrArray(bDibMask()), VarPtr(tSAMask), 4
   
   With tSADest
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lWidth * 4
      .pvData = m_lPtr
   End With
   CopyMemory ByVal VarPtrArray(bDibDest()), VarPtr(tSADest), 4

   xDst = m_tIcon(lIndex).xPosition * 4
   xDstInit = xDst
   xEnd = (m_lIconWidth - 1) * 4
   
   ' Check if all of the colour alpha bits are zero:
   bAllZero = True
   bAllWhite = True
   For y = 0 To m_lIconHeight - 1
      For x = 0 To xEnd Step 4
         If Not (bDibColor(x + 3, y) = 0) Then
            bAllZero = False
         End If
         If (bDibMask(x, y) = 0) Then
            If (bDibColor(x, y) = 255 And bDibColor(x + 1, y) = 255 And bDibColor(x + 2, y) = 255) Then
            ElseIf ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) _
               Or ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
            Else
               bAllWhite = False
            End If
         End If
      Next x
   Next y
   
   For y = 0 To m_lIconHeight - 1
      For x = 0 To xEnd Step 4
      
         If (bDibMask(x, y) > 0) Then
            ' output is transparent
            lR = 0
            lG = 0
            lB = 0
            lAlpha = 0
         Else
            ' output uses the color image:
            If (bAllZero) Then
               lAlpha = 255
            Else
               lAlpha = bDibColor(x + 3, y)
            End If
            
            Select Case eStyle
            Case eIconHighlighted
               If ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) Then
                  lR = lHighR
                  lG = lHighG
                  lB = lHighB
                  lAlpha = lAlpha * 90& \ 255&
               ElseIf ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
                  lR = lHighR
                  lG = lHighG
                  lB = lHighB
                  lAlpha = lAlpha * 90& \ 255&
               Else
                  lR = bDibColor(x + 2, y)
                  lG = bDibColor(x + 1, y)
                  lB = bDibColor(x, y)
               End If
               
            Case eIconNonHighlighted
               lR = bDibColor(x + 2, y)
               lG = bDibColor(x + 1, y)
               lB = bDibColor(x, y)
               If ((bDibColor(x + 2, y) = 221) And (bDibColor(x + 1, y) = 221) And (bDibColor(x, y) = 221)) _
                  Or ((bDibColor(x + 2, y) = 227) And (bDibColor(x + 1, y) = 227) And (bDibColor(x, y) = 227)) Then
                  lAlpha = lAlpha * 50& / 255&
               End If
            
            Case eIconDIsabled
               lR = lResR
               lG = lResG
               lB = lResB
               If (bAllWhite) And (bDibColor(x + 2, y) = 255) And (bDibColor(x + 1, y) = 255) And (bDibColor(x, y) = 255) Then
                  lAlpha = 255
               Else
                  ' Decrease alpha in proportion to the grey value:
                  lGrey = (222& * bDibColor(x + 2, y) + 707& * bDibColor(x + 1, y) + 71& * bDibColor(x, y)) / 1000&
                  lAlpha = lAlpha * (255& - lGrey) / 255&
               End If
               
            End Select
                           
            ' Premultiply alpha:
            lR = lR * lAlpha \ 255
            lG = lG * lAlpha \ 255
            lB = lB * lAlpha \ 255
            
         End If
                  
         bDibDest(xDst, y) = lB
         bDibDest(xDst + 1, y) = lG
         bDibDest(xDst + 2, y) = lR
         bDibDest(xDst + 3, y) = lAlpha
         
         xDst = xDst + 4
      Next x
      
      xDst = xDstInit
      
   Next y

   CopyMemory ByVal VarPtrArray(bDibDest), 0&, 4
   CopyMemory ByVal VarPtrArray(bDibMask), 0&, 4
   CopyMemory ByVal VarPtrArray(bDibColor), 0&, 4


End Sub

Private Sub ChunkResize()
Dim lhWndD As Long
Dim lhDibOld As Long
Dim lPtrOld As Long
Dim lOldWidth As Long
Dim lOldHeight As Long
Dim lHDC As Long
Dim lhDCT As Long
Dim lhBmpOldT As Long

   
   m_lCurrentChunks = m_lCurrentChunks + m_lChunkSize
   
   lhWndD = GetDesktopWindow()
   lHDC = GetDC(lhWndD)
   
   If Not (m_hDC = 0) Then
      ' Cache old DIB
      SelectObject m_hDC, m_hBmpOld
      lhDCT = CreateCompatibleDC(lHDC)
      lhBmpOldT = SelectObject(lhDCT, m_hDib)
      
      lhDibOld = m_hDib
      lPtrOld = m_lPtr
      lOldWidth = m_lWidth
      lOldHeight = m_lHeight
      m_hBmpOld = 0
      m_hDib = 0
      m_lPtr = 0
   Else
      ' Create DC to hold DIB
      m_hDC = CreateCompatibleDC(lHDC)
   End If
   
   ' Create the new DIB
   m_lWidth = m_lCurrentChunks * m_lIconWidth
   m_lHeight = m_lIconHeight
   CreateDIB lHDC, m_lWidth, m_lHeight, m_hDib
   ' Add to DC:
   m_hBmpOld = SelectObject(m_hDC, m_hDib)
   
   If Not (lPtrOld = 0) Then
      ' Copy data from old Dib to new one
      CopyOldDibDataToNew lPtrOld, lOldWidth, lOldHeight
      SelectObject lhDCT, lhBmpOldT
      DeleteObject lhDibOld
      DeleteDC lhDCT
   End If
   
   ReleaseDC lhWndD, lHDC
   
End Sub

Private Sub CopyOldDibDataToNew( _
      ByVal lPtrOld As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
Dim tSAFrom As SAFEARRAY2D
Dim lDibFrom() As Long
Dim tSATo As SAFEARRAY2D
Dim lDibTo() As Long
Dim x As Long
Dim y As Long

   With tSAFrom
      .cbElements = 4
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = lWidth
      .pvData = lPtrOld
   End With
   CopyMemory ByVal VarPtrArray(lDibFrom()), VarPtr(tSAFrom), 4
   
   With tSATo
      .cbElements = 4
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = m_lWidth
      .pvData = m_lPtr
   End With
   CopyMemory ByVal VarPtrArray(lDibTo()), VarPtr(tSATo), 4

   For x = 0 To lWidth - 1
      For y = 0 To lHeight - 1
         lDibTo(x, y) = lDibFrom(x, y)
      Next y
   Next x
   
   CopyMemory ByVal VarPtrArray(lDibTo), 0&, 4
   CopyMemory ByVal VarPtrArray(lDibFrom), 0&, 4

End Sub

Private Function CreateDIB( _
      ByVal lHDC As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByRef hDib As Long _
   ) As Boolean
   m_lPtr = 0
   hDib = 0
   With m_tBI.bmiHeader
      .biSize = Len(m_tBI.bmiHeader)
      .biWidth = lWidth
      .biHeight = lHeight
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = BI_RGB
      .biSizeImage = .biWidth * .biHeight * 4 ' 32 bpp
   End With
   hDib = CreateDIBSection( _
           lHDC, _
           m_tBI, _
           DIB_RGB_COLORS, _
           m_lPtr, _
           0, 0)
   CreateDIB = (hDib <> 0)
   
End Function

Public Sub InitialiseFromInstance( _
      cInstance As cCommandBarImageList _
   )
         
   Destroy
   
   m_hIml = 0
   m_ptrVB6ImageList = 0
   
   m_hIml = cInstance.hIml
   m_ptrVB6ImageList = cInstance.VB6ImageListPtr
   m_lIconWidth = cInstance.IconWidth
   m_lIconHeight = cInstance.IconHeight
      
   ChunkResize
   
End Sub

Public Sub InitialiseFromVariant( _
      ByVal vImageList As Variant _
   )
      
   Destroy
   
   m_hIml = 0
   m_ptrVB6ImageList = 0
   
   Dim hImlFound As Long
   Dim ptrVB6Iml As Long
   
   If (VarType(vImageList) = vbLong) Then
      ' Assume a handle to an image list:
      hImlFound = vImageList
   ElseIf (VarType(vImageList) = vbObject) Then
      ' Assume a VB image list:
      On Error Resume Next
      ' Get the image list initialised..
      vImageList.ListImages(1).Draw 0, 0, 0, 1
      hImlFound = vImageList.hImageList
      If (Err.Number = 0) Then
         ' Check for VB6 image list:
         If (TypeName(vImageList) = "ImageList") Then
            Dim o As Object
            Set o = vImageList
            ptrVB6Iml = ObjPtr(o)
            hImlFound = 0
         End If
      Else
         Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
      End If
      On Error GoTo 0
   End If
    
   If Not (hImlFound = 0) Then
      Dim rc As RECT
      ImageList_GetImageRect hImlFound, 0, rc
      m_lIconWidth = rc.right - rc.left
      m_lIconHeight = rc.bottom - rc.top
      m_hIml = hImlFound
   ElseIf Not (ptrVB6Iml = 0) Then
      m_lIconWidth = vImageList.ImageWidth
      m_lIconHeight = vImageList.ImageHeight
      m_ptrVB6ImageList = ptrVB6Iml
   End If
   
      
   
   ChunkResize
   
End Sub

Public Sub Destroy()
   If Not (m_lPtr = 0) Then
      m_lPtr = 0
   End If
   If Not (m_hBmpOld = 0) Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not (m_hDib = 0) Then
      DeleteObject m_hDib
      m_hDib = 0
   End If
   If Not (m_hDC = 0) Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
   m_lCurrentChunks = 0
   m_iIconCount = 0
   
   Erase m_tIcon
   
End Sub

Private Sub Class_Initialize()
   m_lChunkSize = 10
   m_lAlpha = 64
End Sub

Private Sub Class_Terminate()
   Destroy
End Sub






