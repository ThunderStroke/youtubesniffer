Attribute VB_Name = "modFunc"
Option Explicit
'*******************************
'* Most of the function was not written by me!
'* I took it from the internet
'*******************************
'Time===
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Time End===

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
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
'Load Picture End===

'Terminate Process
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Terminate Process End===

'Round Form
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Round Form End===

'Move rightbottom===
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Move rightbottom End===

'Optimize Usage
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
'Optimize Usage End===

'Remove Menu===
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const MF_REMOVE = &H1000&
Public Const SC_CLOSE = &HF060
'Remove Menu End===

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

Public Function DownHTML(strURL As String, strCharset As String) As String
    
    
    DownHTML = ""
    On Error Resume Next
    Dim objAdoStream As Object
    Set objAdoStream = CreateObject("ADODB.Stream")
    Dim objXMLHTTP As MSXML2.XMLHTTP
    Set objXMLHTTP = New MSXML2.XMLHTTP
    objXMLHTTP.Open "GET", strURL, False
    objXMLHTTP.send
    If objXMLHTTP.readyState = 4 Then
        If strCharset = "" Then
            DownHTML = objXMLHTTP.responseBody
        Else
            objAdoStream.Type = 1
            objAdoStream.Mode = 3
            objAdoStream.Open
            objAdoStream.Write objXMLHTTP.responseBody
            objAdoStream.Position = 0
            objAdoStream.Type = 2
            objAdoStream.Charset = strCharset
            DownHTML = objAdoStream.ReadText
            objAdoStream.Close
        End If
    End If
    Set objAdoStream = Nothing
    Set objXMLHTTP = Nothing
    
End Function

'Public Function UTF8_UrlDecode(ByVal URL As String)
'    Dim B, ub
'    Dim UtfB
'    Dim UtfB1, UtfB2, UtfB3
'    Dim I, n, s
'    n = 0
'    ub = 0
'    For I = 1 To Len(URL)
'        DoEvents
'        B = Mid(URL, I, 1)
'        Select Case B
'        Case "+"
'            s = s & " "
'        Case "%"
'            ub = Mid(URL, I + 1, 2)
'            UtfB = Val("&H" & ub)
'            If UtfB < 128 Then
'                I = I + 2
'                s = s & ChrW(UtfB)
'            Else
'                UtfB1 = (UtfB And &HF) * &H1000
'                UtfB2 = (CInt("&H" & Mid(URL, I + 4, 2)) And &H3F) * &H40
'                UtfB3 = CInt("&H" & Mid(URL, I + 7, 2)) And &H3F
'                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
'                I = I + 8
'            End If
'        Case Else
'            s = s & B
'        End Select
'    Next
'    UTF8_UrlDecode = s
'End Function

Public Function VisitURL(VisitUrlink As String)
    On Error Resume Next
    Shell "rundll32.exe url.dll,FileProtocolHandler " & Chr(34) & VisitUrlink & Chr(34), 1
End Function


Public Sub FitPictureToBox(ByVal picSource As PictureBox, picDestination As PictureBox)
    On Error Resume Next
    Dim nAspectRatio As Single
    Dim nWidth As Single
    Dim nHeight As Single
    
    ' get the original picture's aspect ratio.
    nAspectRatio = picSource.ScaleWidth / picSource.ScaleHeight
    
    ' get the size available.
    nWidth = picDestination.ScaleWidth
    nHeight = picDestination.ScaleHeight
    
    ' Adjust the nWidth/nHeight ratio to match nAspectRatio.
    If nWidth / nHeight > nAspectRatio Then
        ' The area is too short and wide.
        ' Make it narrower.
        nWidth = nAspectRatio * nHeight
    Else
        ' The area is too tall and thin.
        ' Make it shorter.
        nHeight = nWidth / nAspectRatio
    End If
    
    ' Center the image at the correct size.
    picDestination.Cls
    picDestination.PaintPicture picSource.Picture, _
    (picDestination.ScaleWidth - nWidth) / 2, _
    (picDestination.ScaleHeight - nHeight) / 2, _
    nWidth, nHeight
End Sub

Public Function LoadPicture(ByVal strFileName As String) As Picture
    
    Dim IID  As TGUID
    With IID
        .data1 = &H7BF80980
        .data2 = &HBF32
        .data3 = &H101A
        .data4(0) = &H8B
        .data4(1) = &HBB
        .data4(2) = &H0
        .data4(3) = &HAA
        .data4(4) = &H0
        .data4(5) = &H30
        .data4(6) = &HC
        .data4(7) = &HAB
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

Public Sub OptiUsage(ByVal nGetCurrentProcess As Long)
    SetProcessWorkingSetSize nGetCurrentProcess, -1&, -1&
End Sub

Public Sub LoopFormControls(ByVal frmForm As Form)
    Dim objForm As Object
    For Each objForm In frmForm
        DoEvents
    Next
End Sub

Public Sub txtTips(ByVal txtDest As TextBox, ByVal strTips As String, ByVal boolGotFocus As Boolean)
    If boolGotFocus = True Then
        txtDest.Text = strTips
        txtDest.Font.Italic = True
        txtDest.ForeColor = RGB(255 / 2, 255 / 2, 255 / 2)
    Else
        txtDest.Text = ""
        txtDest.Font.Italic = False
        txtDest.ForeColor = vbBlack
    End If
End Sub

