Attribute VB_Name = "modLoadPNG"
Option Explicit
 
Public Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As Long, graphics As Long) As Long
Public Declare Function GdipCreateFromHWND Lib "GDIPlus" (ByVal hwnd As Long, graphics As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As Long) As Long
Public Declare Function GdipGetDC Lib "GDIPlus" (ByVal graphics As Long, hdc As Long) As Long
Public Declare Function GdipReleaseDC Lib "GDIPlus" (ByVal graphics As Long, ByVal hdc As Long) As Long
Public Declare Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Public Declare Function GdipLoadImageFromFile Lib "GDIPlus" (ByVal filename As String, Image As Long) As Long
Public Declare Function GdipCloneImage Lib "GDIPlus" (ByVal Image As Long, cloneImage As Long) As Long
Public Declare Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As Long, Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As Long, Height As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long
Public Declare Function GdipBitmapGetPixel Lib "GDIPlus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, color As Long) As Long
Public Declare Function GdipBitmapSetPixel Lib "GDIPlus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Public Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Public Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
 
 
Public Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal Stream As Long, ByRef Image As Long) As Long

Public Declare Function GdipDrawImage Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single) As Long
 
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112
 
Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
 
Public Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Public Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As Long)
 
Public Enum GpStatus
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum

Public Const ULW_OPAQUE = &H4
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const BI_RGB As Long = 0&
Public Const DIB_RGB_COLORS As Long = 0
Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_STYLE As Long = -16
Public Const GWL_EXSTYLE As Long = -20
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER
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

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)


Public mDC As Long
Public mainBitmap As Long
Public blendFunc32bpp As BLENDFUNCTION
Public token As Long
Public oldBitmap As Long

Public Function LoadResImage(ResId, ByVal ResType As String, ByVal hwnd As Long, ByVal Image As Long)
    Dim picData() As Byte
    Dim picStream As Long
    Dim GpInputbuf As GdiplusStartupInput
    Dim token As Long
    Dim GdiDc As Long
    Dim img As Long
    '---初始化GDI+
    GpInputbuf.GdiplusVersion = 1
    GdiplusStartup token, GpInputbuf
    
    '---转换
    GdipCreateFromHDC Image, GdiDc                                              '从窗体的DC创建一个GdiDc对象,操作Gdidc的时候就会画到窗体上
    picData = LoadResData(ResId, ResType)                                       '读取byte
    CreateStreamOnHGlobal picData(0), 0, picStream                              'byte转成流
    GdipLoadImageFromStream picStream, img                                      '从流加载
    GdipDrawImage GdiDc, img, 0, 0                                              '画到gdidc里，也就是窗体的dc
    
    '--透明
    Dim tempBI As BITMAPINFO
    Dim tempBlend As BLENDFUNCTION
    Dim lngHeight As Long, lngWidth As Long
    Dim curWinLong As Long
    
    Dim graphics As Long
    Dim winSize As Size
    Dim srcPoint As POINTAPI
    
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32
        .biHeight = frmAbout.ScaleHeight
        .biWidth = frmAbout.ScaleWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With
    mDC = CreateCompatibleDC(Image)
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    oldBitmap = SelectObject(mDC, mainBitmap)
    
    Call GdipCreateFromHDC(mDC, graphics)
    
    Call GdipGetImageHeight(img, lngHeight)
    Call GdipGetImageWidth(img, lngWidth)
    Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)
    
    curWinLong = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    SetWindowLong hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    srcPoint.x = 0
    srcPoint.y = 0
    winSize.cx = frmAbout.ScaleWidth
    winSize.cy = frmAbout.ScaleHeight
    
    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    Call GdipDisposeImage(img)
    Call GdipDeleteGraphics(graphics)
    Call UpdateLayeredWindow(hwnd, Image, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
    
    '---释放对象
    GdipDisposeImage img
    GdipDeleteGraphics GdiDc
    GdiplusShutdown token
End Function


