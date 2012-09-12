Attribute VB_Name = "modFormMove"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const WM_CLOSE = &H10

Public pWndProc As Long
Public Const WM_MOVE = &H3


Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Msg = WM_MOVE Then                                                       '?? WM_CLOSE ??????
        Form2.Move frmMain.Left, frmMain.Top + frmMain.Height
        WindowProc = 1
        Exit Function
        
    End If
    
    WindowProc = CallWindowProc(pWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Function WindowProc_Child(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
    If Msg = WM_MOVE Then                                                       '?? WM_CLOSE ??????
        frmMain.Move Form2.Left, Form2.Top - frmMain.Height
        WindowProc_Child = 1
        Exit Function
        
    End If
    
    WindowProc_Child = CallWindowProc(pWndProc, hwnd, Msg, wParam, lParam)
End Function

