Attribute VB_Name = "modFunc"
Option Explicit

Public Function SFProgressBarProc(ByVal This As SFProgressBar, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    SFProgressBarProc = This.WindowProcedure(hWnd, wMsg, wParam, lParam)
End Function
