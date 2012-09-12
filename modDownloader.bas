Attribute VB_Name = "modDownloader"
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE = &HFFFF                                                 '  Infinite timeout
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Const WAIT_TIMEOUT = &H102
Public WaithProcess As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Const LVS_EX_FULLROWSELECT = &H20
Const LVM_FIRST = &H1000
Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'remove menu
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const MF_REMOVE = &H1000&
Public Const SC_CLOSE = &HF060
Public ffplayPID As Long

Public PatcherArg As String

'URLDecode==============================================
Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszURL As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long
Private Const E_POINTER As Long = &H80004003


Function WaitExeRun(cmdIn As String)
    
    Dim dwFlag As Long
    cmdpid = Shell(cmdIn, vbNormalFocus)
    ffplayPID = cmdpid
    On Error Resume Next
    WaithProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, cmdpid)
    While WAIT_TIMEOUT = WaitForSingleObject(WaithProcess, 0)
        DoEvents
    Wend
End Function

' ============================================
' 计算数据
' ============================================
Public Function VBStrFormatByteSize(ByVal FileLengthName As Double) As String
    'On Error Resume Next
    'Dim strSize As String * 128
    'Dim strData As String
    'Dim lPos        As Long
    'StrFormatByteSize LngSize, strSize, 128
    'lPos = InStr(1, strSize, Chr$(0))
    'strData = Left$(strSize, lPos - 1)
    'VBStrFormatByteSize = strData
    
    If FileLengthName < 1024 Then
        VBStrFormatByteSize = FileLengthName & " Bytes"
    ElseIf FileLengthName < 1024 ^ 2 Then
        VBStrFormatByteSize = FormatNumber(FileLengthName / 1024, 2) & " KB"
    ElseIf FileLengthName < 1024 ^ 3 Then
        VBStrFormatByteSize = FormatNumber(FileLengthName / 1024 ^ 2, 2) & " MB"
    ElseIf FileLengthName < 1024 ^ 4 Then
        VBStrFormatByteSize = FormatNumber(FileLengthName / 1024 ^ 3, 2) & " GB"
    ElseIf FileLengthName < 1024 ^ 5 Then
        VBStrFormatByteSize = FormatNumber(FileLengthName / 1024 ^ 4, 2) & " TB"
    Else
        VBStrFormatByteSize = FileLengthName & " Bytes"
    End If
    
End Function

' ============================================
' 根据路径获取文件名
' ============================================
Public Function GetFileNameByPath(ByVal FullPath As String) As String
    On Error Resume Next
    Dim pos, pvs
    FullPath = Replace(FullPath, "/", "\")
    pos = InStrRev(FullPath, "\") + 1
    pvs = InStrRev(FullPath, "?") - pos
    If (pos > 0) Then
        If pvs > 0 Then
            GetFileNameByPath = Mid(FullPath, pos, pvs)
        Else
            GetFileNameByPath = Mid(FullPath, pos)
        End If
    Else
        GetFileNameByPath = FullPath
    End If
End Function
' ============================================
' 根据路径获取扩展名
' ============================================
Public Function GetFileExt(ByVal FullPath)
    On Error Resume Next
    Dim pos, pvs
    pos = InStrRev(FullPath, ".")
    pvs = InStrRev(FullPath, "?") - pos
    If pvs < pos Then pvs = InStrRev(FullPath, "?")
    If pos > 0 Then
        If pvs > 0 Then
            GetFileExt = Mid(FullPath, pos, pvs)
        Else
            GetFileExt = Mid(FullPath, pos)
        End If
    End If
End Function
' ============================================
' 根据路径获取扩展名
' ============================================
Public Function GetFileExt2(ByVal FullPath)
    On Error Resume Next
    Dim pos, pvs
    pos = InStrRev(FullPath, ".")
    pvs = InStr(FullPath, "?")
    If pvs < pos Then pvs = InStrRev(FullPath, "?")
    If pos > 0 Then
        If pvs > pos Then
            GetFileExt2 = Mid(FullPath, pos, pvs - pos)
        Else
            GetFileExt2 = Mid(FullPath, pos)
        End If
    End If
End Function

' ============================================
'   作用 ：创建文件夹
' ============================================
Public Function CheckFolder(strPath) As Boolean
    On Error Resume Next
    Dim Pos1        As Long                                                     '磁盘分隔符
    Dim Pvs1        As Long                                                     '文件分隔符
    Dim strDiver    As String                                                   '磁盘地址
    Dim cfolder     As String                                                   '所有目录
    Dim GetNewsFold As Variant                                                  '目录集合
    Dim fldr        As String                                                   '目录堆栈
    CheckFolder = False
    strPath = Replace(strPath, "/", "\")
    strPath = Replace(strPath, "\\", "\")
    strPath = Replace(strPath, Chr(0), "")
    Pos1 = InStr(strPath, ":")
    If InStr(strPath, ".") > 0 Then
        Pvs1 = InStrRev(strPath, "\")
    Else
        Pvs1 = 0
    End If
    If Pos1 > 0 Then
        strDiver = Mid(strPath, 1, Pos1)
        If Pvs1 > 0 Then
            cfolder = Mid(strPath, Pos1 + 1, Pvs1 - Pos1)
        Else
            cfolder = Mid(strPath, Pos1 + 1, Len(strPath))
        End If
    Else
        Exit Function
    End If
    
    If Left(cfolder, 1) = "\" Then
        cfolder = Right(cfolder, Len(cfolder) - 1)
    Else
        cfolder = cfolder
    End If
    If Right(cfolder, 1) = "/" Then
        cfolder = Left(cfolder, Len(cfolder) - 1)
    Else
        cfolder = cfolder
    End If
    fldr = ""
    If Dir(strDiver & "\" & cfolder, vbDirectory) = "" Then
        GetNewsFold = Split(cfolder, "\")
        Dim i As Long
        For i = 0 To UBound(GetNewsFold)
            If fldr = "" Then
                fldr = GetNewsFold(i)
            Else
                fldr = fldr & "\" & GetNewsFold(i)
            End If
            
            If Dir(strDiver & "\" & fldr, vbDirectory) = "" Then
                MkDir strDiver & "\" & fldr
            End If
        Next
        CheckFolder = True
    Else
        CheckFolder = True
    End If
End Function


Public Sub MoveForm(Movehwnd As Long)
    Call ReleaseCapture
    Call SendMessage(Movehwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Function ConvertToTime(ByVal InTime As Long)
    ConvertToTime = (InTime - InTime Mod 3600) / 3600 & "h:" & ((InTime - InTime Mod 60) / 60) Mod 60 & "m:" & InTime Mod 60 & "s"
End Function



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

'Public Function lvwFullRow(lvwIn As ListView)
    'Dim lStyle As Long
    'lStyle = SendMessage(lvwIn.hwnd, _
    'LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    'lStyle = lStyle Or LVS_EX_FULLROWSELECT
    'Call SendMessage(lvwIn.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
'0, ByVal lStyle)

'End Function
Public Sub SaveToFile(ByVal DirFilename As String, ByRef FileByte() As Byte)
    Open DirFilename For Binary As #3
    Dim x As Double
    For x = 0 To UBound(FileByte)
        Put #3, , FileByte(x)
        DoEvents
    Next
    Close #3
End Sub


'URLDecode==============================================
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
