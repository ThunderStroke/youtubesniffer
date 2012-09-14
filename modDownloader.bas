Attribute VB_Name = "modDownloader"
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'remove menu
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const MF_REMOVE = &H1000&
Public Const SC_CLOSE = &HF060


'Public PatcherArg As String

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

Public Sub SaveToFile(ByVal DirFilename As String, ByRef FileByte() As Byte)
    Open DirFilename For Binary As #3
    Dim X As Double
    For X = 0 To UBound(FileByte)
        Put #3, , FileByte(X)
        DoEvents
    Next
    Close #3
End Sub



