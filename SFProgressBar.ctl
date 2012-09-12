VERSION 5.00
Begin VB.UserControl SFProgressBar 
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ScaleHeight     =   135
   ScaleWidth      =   1575
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   960
      Picture         =   "SFProgressBar.ctx":0000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   200
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "SFProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_stData As SUBCLASSDATA
Private m_rcClient As RECT
Private m_hDCBackBuf As Long
Private m_hBmpBackOld As Long
Private m_hBmpBackNew As Long
Private m_dwMin As Long
Private m_dwMax As Long
Private m_dwValue As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type
Private Type SUBCLASSDATA
    lpOldFunc As Long
    lpNewFunc As Long
    lpVTables(5) As Long
End Type

Private Sub UserControl_Initialize()
    With m_stData
        .lpVTables(0) = &HB82434FF
        .lpVTables(1) = ObjPtr(Me)
        .lpVTables(2) = &H4244489
        .lpVTables(3) = &HB8909090
        .lpVTables(4) = GetAddress(AddressOf SFProgressBarProc)
        .lpVTables(5) = &H9090E0FF
        .lpNewFunc = VarPtr(.lpVTables(0))
    End With
    m_stData.lpOldFunc = SetWindowLong(UserControl.hWnd, -4, m_stData.lpNewFunc)
    m_hDCBackBuf = CreateCompatibleDC(UserControl.hDC)
    m_hBmpBackOld = GetCurrentObject(m_hDCBackBuf, 7)
End Sub

Private Sub UserControl_InitProperties()
        m_dwMax = 100
End Sub

Friend Function WindowProcedure(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg
    Case &H5
        UserControl.Height = 150
        Call GetClientRect(hWnd, m_rcClient)
        If m_hBmpBackNew Then
            Call DeleteObject(m_hBmpBackNew)
        End If
        m_hBmpBackNew = CreateCompatibleBitmap(UserControl.hDC, m_rcClient.Right, _
        m_rcClient.Bottom)
        Call SelectObject(m_hDCBackBuf, m_hBmpBackNew)
    Case &HF
        Dim dwStartPos As Long
        Dim ps As PAINTSTRUCT
        Call BeginPaint(hWnd, ps)
        '*******************************************************************
        '绘制进度条
        '*******************************************************************
        Call BitBlt(m_hDCBackBuf, 0, 0, 3, 10, Picture1.hDC, 0, 0, vbSrcCopy)
        Call StretchBlt(m_hDCBackBuf, 3, 0, m_rcClient.Right - 6, 10, Picture1.hDC, 3, 0, 2, _
10, vbSrcCopy)
        Call BitBlt(m_hDCBackBuf, m_rcClient.Right - 3, 0, 3, 10, Picture1.hDC, 5, 0, vbSrcCopy)
        '*******************************************************************
        '绘制遮挡层
        '*******************************************************************
        dwStartPos = ((m_rcClient.Right - 2) / (m_dwMax - m_dwMin)) * (m_dwValue - m_dwMin) + 1
        Call StretchBlt(m_hDCBackBuf, dwStartPos, 0, m_rcClient.Right - dwStartPos - 1, 10, _
        Picture1.hDC, 8, 0, 2, 10, vbSrcCopy)
        Call BitBlt(ps.hDC, 0, 0, m_rcClient.Right, m_rcClient.Bottom, _
        m_hDCBackBuf, 0, 0, vbSrcCopy)
        Call EndPaint(hWnd, ps)
    End Select
    WindowProcedure = CallWindowProc(m_stData.lpOldFunc, hWnd, wMsg, wParam, lParam)
End Function

Public Property Get Min() As Long
Attribute Min.VB_Description = "返回/设置进度条的最小值。"
Attribute Min.VB_ProcData.VB_Invoke_Property = ";杂项"
       Min = m_dwMin
End Property

Public Property Let Min(ByVal New_Min As Long)
       If New_Min < m_dwMax And New_Min >= 0 Then
          m_dwMin = New_Min
          If m_dwValue < New_Min Then
             m_dwValue = New_Min
          End If
          Call Refresh
       End If
       Call PropertyChanged("Min")
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "返回/设置进度条的最大值。"
Attribute Max.VB_ProcData.VB_Invoke_Property = ";杂项"
       Max = m_dwMax
End Property

Public Property Let Max(ByVal New_Max As Long)
       If New_Max > m_dwMin Then
          m_dwMax = New_Max
          If m_dwValue > New_Max Then
             m_dwValue = New_Max
          End If
          Call Refresh
       End If
       Call PropertyChanged("Max")
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "返回/设置进度条的当前值。"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";外观"
       Value = m_dwValue
End Property

Public Property Let Value(ByVal New_Value As Long)
       If New_Value >= m_dwMin And New_Value <= m_dwMax Then
          m_dwValue = New_Value
          Call Refresh
       End If
       Call PropertyChanged("Value")
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        Call PropBag.WriteProperty("Min", m_dwMin, 0)
        Call PropBag.WriteProperty("Max", m_dwMax, 100)
        Call PropBag.WriteProperty("Value", m_dwValue, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        m_dwMin = PropBag.ReadProperty("Min", 0)
        m_dwMax = PropBag.ReadProperty("Max", 100)
        m_dwValue = PropBag.ReadProperty("Value", 0)
End Sub

Private Function GetAddress(ByVal ptr As Long) As Long
        GetAddress = ptr
End Function

Private Sub UserControl_Terminate()
        Call SelectObject(m_hDCBackBuf, m_hBmpBackOld)
        Call DeleteObject(m_hBmpBackNew)
        Call DeleteDC(m_hDCBackBuf)
        If m_stData.lpNewFunc Then
           Call SetWindowLong(UserControl.hWnd, -4, m_stData.lpOldFunc)
           m_stData.lpNewFunc = 0
       End If
End Sub
