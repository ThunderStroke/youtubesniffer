Attribute VB_Name = "modFunc"
Option Explicit
'Move Form===
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Move Form End===

'URLDecode===
Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszURL As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long
Private Const E_POINTER As Long = &H80004003
'URLDecode End===
Public Function URLDecode( _
    ByVal URL As String, _
    Optional ByVal PlusSpace As Boolean = True) As String
    Dim cchUnescaped As Long
    Dim HRESULT As Long
    If PlusSpace Then URL = Replace$(URL, "+", " ")
    cchUnescaped = Len(URL)
    URLDecode = String$(cchUnescaped, 0)
    HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    If HRESULT = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    End If
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

Public Function ConvertToTime(ByVal InTime As Long)
    ConvertToTime = IIf(((InTime - InTime Mod 3600) / 3600) > 0, (InTime - InTime Mod 3600) / 3600 & " : ", "") & Format(((InTime - InTime Mod 60) / 60) Mod 60, "0#") & " : " & Format(InTime Mod 60, "0#")
End Function

Public Sub MoveForm(Movehwnd As Long)
    Call ReleaseCapture
    Call SendMessage(Movehwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

