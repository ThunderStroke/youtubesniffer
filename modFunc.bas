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

'Load Picture
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
'Load Picture End===

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

Public Function UTF8_UrlDecode(ByVal URL As String)
    Dim B, ub
    Dim UtfB
    Dim UtfB1, UtfB2, UtfB3
    Dim I, n, s
    n = 0
    ub = 0
    For I = 1 To Len(URL)
        DoEvents
        B = Mid(URL, I, 1)
        Select Case B
        Case "+"
            s = s & " "
        Case "%"
            ub = Mid(URL, I + 1, 2)
            UtfB = Val("&H" & ub)
            If UtfB < 128 Then
                I = I + 2
                s = s & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000
                UtfB2 = (CInt("&H" & Mid(URL, I + 4, 2)) And &H3F) * &H40
                UtfB3 = CInt("&H" & Mid(URL, I + 7, 2)) And &H3F
                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                I = I + 8
            End If
        Case Else
            s = s & B
        End Select
    Next
    UTF8_UrlDecode = s
End Function

Public Function VisitURL(VisitUrlink As String)
    On Error Resume Next
    Shell "rundll32.exe url.dll,FileProtocolHandler " & Chr(34) & VisitUrlink & Chr(34), 1
End Function

' Copy the image from pic_src into pic_dst so
' it fits and has the same aspect ratio as the
' original picture.
Public Sub FitPictureToBox(ByVal pic_src As PictureBox, pic_dst As PictureBox)
    On Error Resume Next
    Dim aspect_src As Single
    Dim wid As Single
    Dim hgt As Single
    
    ' get the original picture's aspect ratio.
    aspect_src = pic_src.ScaleWidth / pic_src.ScaleHeight
    
    ' get the size available.
    wid = pic_dst.ScaleWidth
    hgt = pic_dst.ScaleHeight
    
    ' Adjust the wid/hgt ratio to match aspect_src.
    If wid / hgt > aspect_src Then
        ' The area is too short and wide.
        ' Make it narrower.
        wid = aspect_src * hgt
    Else
        ' The area is too tall and thin.
        ' Make it shorter.
        hgt = wid / aspect_src
    End If
    
    ' Center the image at the correct size.
    pic_dst.Cls
    pic_dst.PaintPicture pic_src.Picture, _
    (pic_dst.ScaleWidth - wid) / 2, _
    (pic_dst.ScaleHeight - hgt) / 2, _
    wid, hgt
End Sub

Public Function LoadPicture(ByVal strFileName As String) As Picture
    
    Dim IID  As TGUID
    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    On Error GoTo ERR_LINE
    OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
    Exit Function
ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)
End Function


Public Function ExtractMatch(Text, Pattern)
    Dim Regex As Object
    Dim Matches
    
    Set Regex = New RegExp
    Regex.Pattern = Pattern
    
    Set Matches = Regex.Execute(Text)
    If Matches.Count = 0 Then
        ExtractMatch = ""
        Exit Function
    End If
    
    ExtractMatch = Matches(0).SubMatches(0)
End Function


