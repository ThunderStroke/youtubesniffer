Attribute VB_Name = "modStickForm"
'运行后,移动Form2到frmmain附近,例其自动吸附
'当前只能使Form2吸附到frmmain的右边或下边
'可添加其它代码使可吸附到其它位置
'吸附后,试着移动或缩放frmmain
'
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public PrevWndProc_frmmain As Long
Public PrevWndProc_Form2 As Long

Public Const GWL_WNDPROC = (-4)
Private Const WM_MOVE = &H3
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SIZE = &H5

Private isRight As Boolean                                                      '标志,Form2是否吸附在frmmain的右边
Private isBottom As Boolean                                                     '标志,Form2是否吸附在frmmain的底边

Dim oldX&, oldY&                                                                '保存frmmain移动前与From2之间的距离
Const Distance As Integer = 200                                                 '吸附距离,小于此数值时将自动吸附

Public Function SubWndProc_frmmain(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Msg = WM_MOVE Then                                                       '当frmmain移动时
        
        If isRight Then                                                         '如果Form2吸附在右边
            Form2.Left = frmMain.Left + frmMain.Width
            Form2.Top = frmMain.Top - oldY
        End If
        
        If isBottom Then                                                        '如果Form2吸附在下边
            Form2.Top = frmMain.Top + frmMain.Height
            Form2.Left = frmMain.Left - oldX
        End If
    ElseIf Msg = WM_NCLBUTTONDOWN Then                                          '当点击frmmain的标题栏时,保存Form2与frmmain的距离
        oldX = frmMain.Left - Form2.Left
        oldY = frmMain.Top - Form2.Top
    ElseIf Msg = WM_SIZE Then                                                   '当缩放frmmain时
        If isRight Then
            Form2.Left = frmMain.Left + frmMain.Width
        End If
        
        If isBottom Then
            Form2.Top = frmMain.Top + frmMain.Height
        End If
    End If
    SubWndProc_frmmain = CallWindowProc(PrevWndProc_frmmain, hwnd, Msg, wParam, lParam)
End Function

Public Function SubWndProc_Form2(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Msg = WM_MOVE Then                                                       '当Form2移动时
        '如果Form2的顶部与From1的底部距离小于设定值,则自动吸附
        If Abs(Form2.Top - (frmMain.Top + frmMain.Height)) < Distance Then
            Form2.Top = frmMain.Top + frmMain.Height
            isBottom = True
        Else
            isBottom = False
        End If
        '如果Form2的左边与From1的右边距离小于设定值,则自动吸附
        If Abs(Form2.Left - (frmMain.Left + frmMain.Width)) < Distance Then
            Form2.Left = frmMain.Left + frmMain.Width
            isRight = True
        Else
            isRight = False
        End If
    End If
    SubWndProc_Form2 = CallWindowProc(PrevWndProc_Form2, hwnd, Msg, wParam, lParam)
End Function





