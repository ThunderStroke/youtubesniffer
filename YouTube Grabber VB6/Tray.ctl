VERSION 5.00
Begin VB.UserControl Tray 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
                                                                 
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIM_ADD = &H0                                                     '在任务栏中增加一个图标
Private Const NIM_DELETE = &H2                                                  '删除任务栏中的一个图标
Private Const NIM_MODIFY = &H1                                                  '修改任务栏中个图标信息
Private Const NIM_SETFOCUS = &H3
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
    cbSize        As Long                                                       '该数据结构的大小
    hwnd          As Long                                                       '处理任务栏中图标的窗口句柄
    uID           As Long                                                       '定义的任务栏中图标的标识
    uFlags        As Long                                                       '任务栏图标功能控制，可以是以下值的组合（一般全包括）
    uCallbackMessage As Long                                                    '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
    hIcon         As Long                                                       '任务栏中的图标的控制句柄
    szTip         As String * 128                                               '图标的提示信息。若要产生气泡提示信息，则一定要128才性，为64则无法生成气泡，其它功能都正常，原因不明
    dwState       As Long
    dwStateMask   As Long
    szInfo        As String * 256                                               '气泡提示内容
    uTimeout      As Long                                                       '气泡提示显示时间
    szInfoTitle   As String * 64                                                '气泡提示标题
    dwInfoFlags   As Long                                                       '气泡提示类型，见 NIIF_*** 部分
End Type
Public Enum Ico                                                                 '气泡提示类型
    NIIF_NONE = &H0                                                             '无图标
    NIIF_INFO = &H1                                                             '信息图标
    NIIF_WARNING = &H2                                                          '警告图标
    NIIF_ERROR = &H3                                                            '错误图标
    NIIF_GUID = &H4                                                             '托盘的图标
End Enum
Private IconData  As NOTIFYICONDATA
Dim PictureIcon   As StdPicture                                                 '图标
Event PopupMenu()                                                               '右键菜单事件
Event Click()                                                                   '托盘单击事件

'创建托盘
Public Sub Show()
    '生成系统托盘图标
    With IconData
        .cbSize = Len(IconData)
        .hwnd = UserControl.hwnd                                                '.Extender.Parent.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE                                        '响应鼠标事件 'WM_LBUTTONDOWN
        If PictureIcon Is Nothing Then
            .hIcon = UserControl.Extender.Parent.Icon                           '默认为窗口的图标
        Else
            .hIcon = PictureIcon
        End If
    End With
    Shell_NotifyIcon NIM_ADD, IconData                                          '增加托盘图标
End Sub

'\\\\\气泡提示
Public Sub ShowBubble(ByVal Title As String, Text As String, Optional ByVal Style As Ico)
    With IconData
        .szInfoTitle = Title & Chr(0)
        .szInfo = Text & Chr(0)
        .dwInfoFlags = Style
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData                                       '修改托盘图标及相关信息
End Sub

'更换托盘图标
Public Sub ChangeIcon(Optional ByVal Icon = 0)
    With IconData
        .szInfoTitle = Chr(0)
        .szInfo = Chr(0)
        .hIcon = Icon
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData                                       '更换托盘图标
End Sub

'隐藏托盘
Public Sub Hide()
    Shell_NotifyIcon NIM_DELETE, IconData                                       '卸载托盘图标
End Sub

'程序关闭 控件销毁时自动清除残留的图标
Private Sub UserControl_Terminate()
    Call Hide
End Sub

Property Set Picture(New_Picture As StdPicture)                                 '托盘的图片
    Set PictureIcon = New_Picture
    Set UserControl.Picture = New_Picture
End Property

Property Get Picture() As StdPicture                                            '控件背景图
    Set Picture = PictureIcon
End Property

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set PictureIcon = PropBag.ReadProperty("PictureIcon", Nothing)
    If PictureIcon Is Nothing Then Else Set UserControl.Picture = PictureIcon
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 32 * 15
    UserControl.Height = 32 * 15
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PictureIcon", PictureIcon, Nothing)
End Sub

'获取托盘的事件
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Screen.TwipsPerPixelX <> 15 Then Exit Sub
    If Button = 1 Then RaiseEvent Click: Exit Sub
    If Button = 2 Then RaiseEvent PopupMenu: Exit Sub
End Sub
                                                                         
