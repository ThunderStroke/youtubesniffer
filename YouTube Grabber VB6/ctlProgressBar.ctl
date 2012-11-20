VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   CanGetFocus     =   0   'False
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ScaleHeight     =   9
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ToolboxBitmap   =   "ctlProgressBar.ctx":0000
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================
'文 件 名: SFProgressBar.ctl
'当前版本: 1.0
'注    释: 随风进度条控件
'公    司: 随风工作室
'开发人员: 随风の追梦
'创建时间: 2010-5-29 12:02
'修改时间: ----
'完成时间: 2010-5-29 12:24
'邮    箱: 497759649@QQ.COM

'版    本: 1.0
'开发时间: 2010-5-29 12:02 ~ 2010-5-29 12:24
'开发人员: 随风の追梦
'描    述: 一个精简版的进度条控件

'本控件切勿用作商业用途，作者不承担因此造成的任何损失。
'如需转载，请保留作者信息。
'======================================================================================
Option Explicit
'API函数声明
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Type RECT
        Left            As Long
        Top             As Long
        Right           As Long
        Bottom          As Long
End Type
Private m_hDCOffScreen As Long
Private m_hBmpOffOld As Long
Private m_BmpBar As StdPicture
Private m_dwMin As Long
Private m_dwMax As Long
Private m_dwValue As Long
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_Initialize
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
        m_hDCOffScreen = CreateCompatibleDC(UserControl.hDC)                            '创建离屏设备环境
        Set m_BmpBar = LoadResPicture(1000, vbResBitmap)
        m_hBmpOffOld = SelectObject(m_hDCOffScreen, m_BmpBar.Handle)                    '选入进度条位图
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_InitProperties
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
        m_dwMax = 100
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_Paint
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_Paint()
        Dim lpRectClient As RECT
        Dim dwStartPos As Long
        Call GetClientRect(UserControl.hWnd, lpRectClient)                              '获取窗口客户区矩形
        '******************************************************************************
        '绘制进度条
        '******************************************************************************
        Call BitBlt(UserControl.hDC, 0, 0, 3, 10, m_hDCOffScreen, 0, 0, vbSrcCopy)
        Call StretchBlt(UserControl.hDC, 3, 0, lpRectClient.Right - 6, 10, m_hDCOffScreen, 3, 0, 2, _
        10, vbSrcCopy)
        Call BitBlt(UserControl.hDC, lpRectClient.Right - 3, 0, 3, 10, m_hDCOffScreen, 5, 0, vbSrcCopy)
        '******************************************************************************
        '绘制遮挡层
        '******************************************************************************
        dwStartPos = ((lpRectClient.Right - 2) / (m_dwMax - m_dwMin)) * (m_dwValue - m_dwMin) + 1
        Call StretchBlt(UserControl.hDC, dwStartPos, 0, lpRectClient.Right - dwStartPos - 1, 10, _
        m_hDCOffScreen, 8, 0, 2, 10, vbSrcCopy)
End Sub
'--------------------------------------------------------------------------------------
'属 性 名: Min
'描    述: 返回/设置进度条的最小值
'--------------------------------------------------------------------------------------
Public Property Get Min() As Long
Attribute Min.VB_Description = "返回/设置进度条的最小值。"
Attribute Min.VB_ProcData.VB_Invoke_Property = ";杂项"
       Min = m_dwMin
End Property

Public Property Let Min(ByVal New_Min As Long)
       If New_Min >= 0 And New_Min < m_dwMax Then
          m_dwMin = New_Min
          If m_dwValue < New_Min Then
             m_dwValue = New_Min
          End If
          Call Refresh
       End If
End Property
'--------------------------------------------------------------------------------------
'属 性 名: Max
'描    述: 返回/设置进度条的最大值
'--------------------------------------------------------------------------------------
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
End Property
'--------------------------------------------------------------------------------------
'属 性 名: Value
'描    述: 返回/设置进度条的当前值
'--------------------------------------------------------------------------------------
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
End Property
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_Resize
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_Resize()
        UserControl.Height = 150
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_WriteProperties
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        With PropBag
             Call .WriteProperty("Min", m_dwMin, 0)
             Call .WriteProperty("Max", m_dwMax, 100)
             Call .WriteProperty("Value", m_dwValue, 0)
        End With
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_ReadProperties
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        With PropBag
             m_dwMin = .ReadProperty("Min", 0)
             m_dwMax = .ReadProperty("Max", 100)
             m_dwValue = .ReadProperty("Value", 0)
        End With
End Sub
'--------------------------------------------------------------------------------------
'事 件 名: UserControl_Terminate
'描    述: ----
'--------------------------------------------------------------------------------------
Private Sub UserControl_Terminate()
        Call SelectObject(m_hDCOffScreen, m_hBmpOffOld)                                 '选入原离屏位图
        Set m_BmpBar = Nothing
        Call DeleteDC(m_hDCOffScreen)                                                   '删除离屏设备环境
End Sub
