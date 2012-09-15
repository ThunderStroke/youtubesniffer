Attribute VB_Name = "modFunc"
Option Explicit
'Move Form===
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Move Form End===

'URLDecode===
Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" (ByVal pszURL As String, ByVal pszUnescaped As String, ByRef pcchUnescaped As Long, ByVal dwFlags As Long) As Long
Private Const E_POINTER As Long = &H80004003
'URLDecode End===

'Utf to Unicode===
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
'Utf to Unicode End===

Public Function URLDecode(ByVal URL As String, Optional ByVal PlusSpace As Boolean = True) As String
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

Public Function Utf8ToUnicode(ByRef Utf() As Byte) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    lLength = UBound(Utf) - LBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    Utf8ToUnicode = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Utf8ToUnicode), lBufferSize)
    If lRet <> 0 Then
        Utf8ToUnicode = Left(Utf8ToUnicode, lRet)
    Else
        Utf8ToUnicode = ""
    End If
End Function

Public Function DownHTML(strURL As String, strCharset As String)
    DownHTML = ""
    On Error Resume Next
    Dim Ado_Stream As Object
    Set Ado_Stream = CreateObject("ADODB.Stream")
    Dim Obj_XMLHTTP As MSXML2.XMLHTTP
    Set Obj_XMLHTTP = New MSXML2.XMLHTTP
    Obj_XMLHTTP.Open "get", strURL, False
    Obj_XMLHTTP.send
    If Obj_XMLHTTP.readyState = 4 Then
        If strCharset = "" Then
            DownHTML = Obj_XMLHTTP.responseBody
        Else
            Ado_Stream.Type = 1
            Ado_Stream.Mode = 3
            Ado_Stream.Open
            Ado_Stream.Write Obj_XMLHTTP.responseBody
            Ado_Stream.Position = 0
            Ado_Stream.Type = 2
            Ado_Stream.Charset = strCharset
            DownHTML = Ado_Stream.ReadText
            Ado_Stream.Close
        End If
    End If
    Set Ado_Stream = Nothing
    Set Obj_XMLHTTP = Nothing
End Function

