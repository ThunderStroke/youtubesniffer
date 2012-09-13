VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WininetDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" _
    (ByVal lpszCallerName As String, _
    ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, _
    ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
    (ByVal hOpen As Long, _
    ByVal sUrl As String, _
    ByVal sHeaders As String, _
    ByVal lLength As Long, _
    ByVal lFlags As Long, _
    ByVal lContext As Long) _
    As Long

Private Declare Function InternetConnect Lib "wininet.dll" _
    Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, _
    ByVal lpszServerName As String, _
    ByVal nProxyPort As Integer, _
    ByVal lpszUsername As String, _
    ByVal lpszPassword As String, _
    ByVal dwService As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As Long, _
    ByRef sBuffer As Byte, _
    ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) _
    As Integer
    
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
    (ByVal hHttpRequest As Long, _
    ByVal lInfoLevel As Long, _
    ByVal sBuffer As Any, _
    ByRef lBufferLength As Long, _
    ByRef lIndex As Long) As Integer

Private Declare Function HttpOpenRequest Lib "wininet.dll" _
    Alias "HttpOpenRequestA" _
    (ByVal hInternetSession As Long, _
    ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, _
    ByVal lpszVersion As String, _
    ByVal lpszReferer As String, _
    ByVal lpszAcceptTypes As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" _
    Alias "HttpSendRequestA" _
    (ByVal hHttpRequest As Long, _
    ByVal sHeaders As String, _
    ByVal lHeadersLength As Long, _
    ByVal sOptional As String, _
    ByVal lOptionalLength As Long) As Boolean
    
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" _
    Alias "HttpAddRequestHeadersA" _
    (ByVal hHttpRequest As Long, _
    ByVal sHeaders As String, _
    ByVal lHeadersLength As Long, _
    ByVal lModifiers As Long) As Integer
    
Private Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" _
    (ByVal lpszUrlName As String, _
    ByVal lpszCookieName As String, _
    ByVal lpszCookieData As String) _
    As Boolean

Private Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" _
    (ByVal lpszUrlName As String, _
    ByVal lpszCookieName As String, _
    ByVal lpszCookieData As String, _
    lpdwSize As Long) _
    As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInternetHandle As Long) As Boolean
    
Private Declare Function DeleteUrlCacheEntry Lib "wininet" _
    Alias "DeleteUrlCacheEntryA" _
    (ByVal lpszUrlName As String) As Long
    
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
    
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
    
Private Const GWL_WNDPROC = (-4)
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3&
Private Const INVALID_HANDLE_VALUE = -1&
Private Const CREATE_ALWAYS = 2
Private Const FILE_BEGIN = 0
Private Const OPEN_ALWAYS = 4

Public Enum wd_Status
    s1 = 1                                                                      ' 开始会话
    s2 = 2                                                                      ' 连接服务器
    s3 = 3                                                                      ' 发送请求
    s4 = 4                                                                      ' 请求成功
    s5 = 5                                                                      ' 开始接收数据
    s6 = 6                                                                      ' 下载结束
End Enum

Public Enum wd_ErrorCode
    e1 = 1                                                                      ' 会话失败
    e2 = 2                                                                      ' 连接服务器失败
    e3 = 3                                                                      ' 发送请求失败
    e4 = 4                                                                      ' 下载被中止
    e5 = 5                                                                      ' 文件已经存在
    e6 = 6                                                                      ' 创建目录出错
    e7 = 7                                                                      ' 打开文件时出错
End Enum
    
Private m_Method                As String
Private m_URL                   As String
Private m_Referer               As String
Private m_UserAgent             As String
Private m_Cookie                As String
Private m_Charset               As String
Private m_Pakesize              As Long
Private m_ReDown                As Boolean
Private m_Filename              As String
Private m_PostData              As String
Private m_IsUpload              As Boolean

Private server                  As String
Private port                    As Long
Private script                  As String

Private blnDownLoad             As Boolean                                      '是否下载中

Private hInternetOpen           As Long
Private hInternetConnect        As Long
Private hHttpOpenRequest        As Long

Private Buffer                  As String
Private buflen                  As Long
Private RetQueryInfo            As Boolean

Private dtTimerStart            As Double                                       '下载时间,用来计算速度
Private TimeConsuming           As Double                                       '下载耗时
Private dbSpeed                 As Double                                       '下载速度

'下载库通过该事件通知文件下载进度
Event Progress(ByVal dbFileSize As Double, ByVal dbFinished As Double, ByVal dbSpeed As Double)
'下载库通过该事件通知下载过程结束
Event DownloadEnd(ByVal bSuccess As Boolean)
'下载库通过该函数通知下载状态
Event Status(ByVal lStatus As wd_Status)
'下载库通过该函数通知下载过程中发生的错误
Event Error(ByVal lErrorCode As wd_ErrorCode)
            
' 返回版本信息
Public Property Get Version()
    Version = "Pentium WininetDown 1.0.0"
End Property

'类初始化
Private Sub Class_Initialize()
    m_Method = "GET"
    m_URL = ""
    m_Referer = ""
    m_UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
    m_Cookie = ""
    m_Charset = "gb2312"
    m_Pakesize = 1024
    m_ReDown = True
    m_Filename = ""
    m_PostData = ""
    m_IsUpload = False
    
    server = ""
    port = 80
    script = ""
    
    blnDownLoad = False
    
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0
End Sub

'类注销
Private Sub Class_Terminate()
    blnDownLoad = False
End Sub

'Method
Public Property Let Method(ByVal oValue As String)
    m_Method = oValue
End Property
Public Property Get Method() As String
    Method = m_Method
End Property

'URL
Public Property Let URL(ByVal oValue As String)
    m_URL = oValue
End Property
Public Property Get URL() As String
    URL = m_URL
End Property

'Referer
Public Property Let Referer(ByVal oValue As String)
    m_Referer = oValue
End Property
Public Property Get Referer() As String
    Referer = m_Referer
End Property

'UserAgent
Public Property Let UserAgent(ByVal oValue As String)
    m_UserAgent = oValue
End Property
Public Property Get UserAgent() As String
    UserAgent = m_UserAgent
End Property

'Cookie
Public Property Let Cookie(ByVal oValue As String)
    m_Cookie = oValue
End Property
Public Property Get Cookie() As String
    Cookie = m_Cookie
End Property

'设置返回字符串类型
Public Property Let Charset(ByVal oValue As String)
    m_Charset = oValue
End Property
Public Property Get Charset() As String
    Charset = m_Charset
End Property

'Pakesize
Public Property Let Pakesize(ByVal oValue As Long)
    m_Pakesize = oValue
End Property
Public Property Get Pakesize() As Long
    Pakesize = m_Pakesize
End Property

'ReDown
Public Property Let ReDown(ByVal oValue As Boolean)
    m_ReDown = oValue
End Property
Public Property Get ReDown() As Boolean
    ReDown = m_ReDown
End Property

'Filename
Public Property Let FileName(ByVal oValue As String)
    m_Filename = oValue
End Property
Public Property Get FileName() As String
    FileName = m_Filename
End Property

'PostData
Public Property Let PostData(ByVal oValue As String)
    m_PostData = oValue
End Property
Public Property Get PostData() As String
    PostData = m_PostData
End Property

' IsUpload
Public Property Let IsUpload(ByVal oValue As Boolean)
    m_IsUpload = oValue
End Property
Public Property Get IsUpload() As Boolean
    IsUpload = m_IsUpload
End Property

' SplitCode
Public Property Let SplitCode(ByVal oValue As String)
    m_SplitCode = oValue
End Property
Public Property Get SplitCode() As String
    SplitCode = m_SplitCode
End Property

' 使用Adodb.Stream处理二进制数据
Private Function BytesToBstr(ByVal strBody, ByVal CodeBase)
    Dim Ados
    Set Ados = CreateObject("Adodb.Stream")
    Ados.Type = 1
    Ados.Mode = 3
    Ados.Open
    Ados.Write strBody
    Ados.Position = 0
    Ados.Type = 2
    Ados.Charset = CodeBase
    BytesToBstr = Ados.ReadText
    Ados.Close
    Set Ados = Nothing
End Function

' 分析URL
Private Sub URLAnalysis()
    Dim bStart As Long
    If m_URL <> "" Then
        server = Replace(m_URL, "http://", "", 1, -1, vbTextCompare)
        bStart = InStr(server, "/")
        If bStart <= 0 Then
            script = "/"
            server = server
        Else
            script = Mid(server, bStart)
            server = Mid(server, 1, bStart - 1)
        End If
        If InStr(server, ":") > 0 Then
            Dim osvr As Variant
            osvr = Split(server, ":")
            server = osvr(0)
            port = CLng(osvr(1))
        Else
            server = server
            port = 80
        End If
    End If
End Sub

' 执行下载
Public Function Download() As Boolean
    blnDownLoad = True
    Dim bRet As Boolean
    ' 分析URL
    URLAnalysis
    ' 建立一个internet会话
    RaiseEvent Status(s1)
    hInternetOpen = InternetOpen(m_UserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hInternetOpen = 0 Then
        Download = False
        RaiseEvent Error(e1)
        Exit Function
    End If
    
    ' 连接服务器
    RaiseEvent Status(s2)
    hInternetConnect = InternetConnect(hInternetOpen, server, port, vbNullString, "HTTP/1.0", INTERNET_SERVICE_HTTP, 0, 0)
    If hInternetConnect = 0 Then
        Download = False
        RaiseEvent Error(e2)
        Exit Function
    End If
    
    ' 发送请求下载指定文件
    RaiseEvent Status(s3)
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, m_Method, script, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If hHttpOpenRequest = 0 Then
        Download = False
        RaiseEvent Error(e3)
        Exit Function
    End If
    ' 请求成功
    Download = True
    RaiseEvent Status(s4)
    
    Dim sHeader As String
    If m_Referer <> "" Then
        sHeader = sHeader & "Referer: " & m_Referer & vbCrLf
    End If
    If UCase(m_Method) = "POST" Then
        sHeader = sHeader & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    End If
    ' 发送客户端信息
    bRet = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
    ' 发送Cookie
    If m_Cookie <> "" Then
        InternetSetCookie m_URL, vbNullString, m_Cookie
    End If
    ' 发送数据
    'If (UCase(m_Method) = "POST") And (m_PostData <> "") Then
    Dim lpszPostData As String
    Dim lPostDataLen As Long
    
    lpszPostData = m_PostData
    lPostDataLen = Len(lpszPostData)
    bRet = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, lpszPostData, lPostDataLen)
    'End If
    
    Buffer = Space$(1024)
    buflen = 1024
    RetQueryInfo = HttpQueryInfo(hHttpOpenRequest, 21, Buffer, buflen, 0)
    If RetQueryInfo Then
        Buffer = Mid$(Buffer, 1, buflen)
        'Debug.Print Buffer
    Else
        Buffer = ""
    End If
End Function

'退出下载
Public Sub bStop()
    blnDownLoad = False
    InternetCloseHandle hHttpOpenRequest
    InternetCloseHandle hInternetConnect
    InternetCloseHandle hInternetOpen
    
    DeleteUrlCacheEntry m_URL
End Sub

'取得文件头
Public Function GetHeader(Optional hdrName As String = vbNullString) As String
    On Error Resume Next
    Dim tmp As Long
    Dim tmp2 As String
    If blnDownLoad = False Then
        GetHeader = "0"
        bStop
        Exit Function
    End If
    If Buffer <> "" Then
        Select Case UCase$(hdrName)
        Case "CONTENT-LENGTH"
            tmp = InStr(Buffer, "Content-Length")
            tmp2 = Mid$(Buffer, tmp + 16, Len(Buffer))
            tmp = InStr(tmp2, Chr$(0))
            GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
        Case "CONTENT-TYPE"
            tmp = InStr(Buffer, "Content-Type")
            tmp2 = Mid$(Buffer, tmp + 14, Len(Buffer))
            tmp = InStr(tmp2, Chr$(0))
            GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
        Case "DATE"
            tmp = InStr(Buffer, "Date")
            tmp2 = Mid$(Buffer, tmp + 6, Len(Buffer))
            tmp = InStr(tmp2, Chr$(0))
            GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
        Case "LAST-MODIFIED"
            tmp = InStr(Buffer, "Last-Modified")
            tmp2 = Mid$(Buffer, tmp + 15, Len(Buffer))
            tmp = InStr(tmp2, Chr$(0))
            GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
        Case "SERVER"
            tmp = InStr(Buffer, "Server")
            tmp2 = Mid$(Buffer, tmp + 8, Len(Buffer))
            tmp = InStr(tmp2, Chr$(0))
            GetHeader = CStr(Mid$(tmp2, 1, tmp - 1))
        Case vbNullString
            GetHeader = Buffer
        Case Else
            GetHeader = "0"
        End Select
    Else
        GetHeader = "0"
    End If
End Function

' 取得文件大小
Public Function FileSize() As Long
    On Error Resume Next
    FileSize = GetHeader("Content-Length")
    If Err Then
        FileSize = 1
    End If
End Function

' 取得Cookie
Public Function GetCookie(Optional ByVal CookName As String = vbNullString) As String
    On Error Resume Next
    Dim sCookieVal As String * 256
    Dim bRet As Boolean
    bRet = InternetGetCookie(m_URL, CookName, sCookieVal, 255)
    If bRet = False Then
        GetCookie = ""
    Else
        sCookieVal = Replace(sCookieVal, Chr(0), "")
        sCookieVal = Replace(sCookieVal, Chr(10), "")
        sCookieVal = Replace(sCookieVal, Chr(13), "")
        GetCookie = sCookieVal
    End If
End Function

'取得文本
Public Function DownLoadToText() As String
    On Error Resume Next
    If Download = False Then
        DownLoadToText = ""
        Exit Function
    End If
    Dim TempBody            As String
    Dim sReadBuffer()       As Byte
    Dim TatolSize           As Double
    Dim lNumberOfBytesRead  As Long
    Dim ReadTatols          As Double
    Dim bRet                As Boolean
    
    DownLoadToText = ""
    
    If blnDownLoad = False Then
        ' 下载被中止
        bStop
        DownLoadToText = ""
        RaiseEvent Error(e4)
        RaiseEvent DownloadEnd(False)
        Exit Function
    End If
    
    ' 取得文件大小
    TatolSize = FileSize
    ' 时间开始
    dtTimerStart = Timer()
    Do While True
        ReDim sReadBuffer(m_Pakesize)
        bRet = InternetReadFile(hHttpOpenRequest, sReadBuffer(0), m_Pakesize, lNumberOfBytesRead)
        If lNumberOfBytesRead > 0 And bRet = True Then
            ReDim Preserve sReadBuffer(0 To lNumberOfBytesRead - 1)
            TempBody = TempBody & BytesToBstr(sReadBuffer, m_Charset)
            ReadTatols = ReadTatols + lNumberOfBytesRead
            TimeConsuming = Timer() - dtTimerStart
            If TimeConsuming > 0 Then
                dbSpeed = ReadTatols / TimeConsuming
            End If
            ' 开始接收数据
            RaiseEvent Status(s5)
            ' 进度信息
            RaiseEvent Progress(TatolSize, ReadTatols, dbSpeed)
        Else
            Exit Do
        End If
        
        DoEvents
        If blnDownLoad = False Then
            ' 下载被中止
            bStop
            DownLoadToText = ""
            RaiseEvent Error(e4)
            RaiseEvent DownloadEnd(False)
            Exit Do
        End If
    Loop
    '取得文本
    DownLoadToText = TempBody
    
    If blnDownLoad = True Then
        '下载完成
        bStop
        RaiseEvent Status(s6)
        RaiseEvent DownloadEnd(True)
    End If
End Function

'下载到文件
Public Function DownloadToFile() As Boolean
    On Error Resume Next
    If Download = False Then
        DownloadToFile = False
        Exit Function
    End If
    Dim sReadBuffer()       As Byte
    Dim TatolSize           As Double
    Dim lNumberOfBytesRead  As Long
    Dim ReadTatols          As Double
    Dim bRet                As Boolean
    
    Dim FileHandle          As Long                                             '文件操作句柄
    
    If blnDownLoad = False Then
        ' 下载被中止
        bStop
        DownloadToFile = False
        RaiseEvent Error(e4)
        RaiseEvent DownloadEnd(False)
        Exit Function
    End If
    
    ' 取得文件大小
    TatolSize = FileSize
    
    '文件存在
    If Len(Dir$(m_Filename)) = TatolSize Then
        If m_ReDown = False Then
            ' 文件已经存在
            bStop
            DownloadToFile = False
            RaiseEvent Error(e5)
            RaiseEvent DownloadEnd(False)
            Exit Function
        Else
            ' 删除
            Kill FileName
        End If
    Else
        ' 删除
        Kill FileName
    End If
    
    If CheckFolder(m_Filename) = True Then
        FileHandle = CreateFile(m_Filename, GENERIC_WRITE, 0, ByVal 0&, OPEN_ALWAYS, 0, ByVal 0&)
    Else
        ' 创建目录出错
        bStop
        DownloadToFile = False
        RaiseEvent Error(e6)
        RaiseEvent DownloadEnd(False)
        Exit Function
    End If
    If Err Then
        Err.Clear
        '关闭文件
        CloseHandle FileHandle
        ' 打开文件时出错
        bStop
        DownloadToFile = False
        RaiseEvent Error(e7)
        RaiseEvent DownloadEnd(False)
        Exit Function
    End If
    '设定写入位置
    SetFilePointer FileHandle, 0, ByVal 0&, FILE_BEGIN
    ' 时间开始
    dtTimerStart = Timer()
    Do While True
        ReDim sReadBuffer(m_Pakesize)
        bRet = InternetReadFile(hHttpOpenRequest, sReadBuffer(0), m_Pakesize, lNumberOfBytesRead)
        If lNumberOfBytesRead > 0 And bRet = True Then
            ReDim Preserve sReadBuffer(0 To lNumberOfBytesRead - 1)
            ReadTatols = ReadTatols + lNumberOfBytesRead
            Dim ret As Long
            WriteFile FileHandle, sReadBuffer(0), lNumberOfBytesRead, ret, ByVal 0&
            TimeConsuming = Timer() - dtTimerStart
            If TimeConsuming > 0 Then
                dbSpeed = ReadTatols / TimeConsuming
            End If
            ' 开始接收数据
            RaiseEvent Status(s5)
            ' 进度信息
            RaiseEvent Progress(TatolSize, ReadTatols, dbSpeed)
            'frmDownload.lblTime = Format(CDate(Int(Timer())) - CDate(Int(dtTimerStart)), "hh:mm:ss")
            If frmDownload.Visible = True Then
                Dim SecTick As Long
                SecTick = Int(Timer() - dtTimerStart)
                frmDownload.lblTime = IIf(SecTick \ 3600 >= 10, SecTick \ 3600 & "h", 0 & SecTick \ 3600 & "h") & ":" & IIf((SecTick Mod 3600) \ 60 >= 10, (SecTick Mod 3600) \ 60 & "m", "0" & (SecTick Mod 3600) \ 60 & "m") & ":" & IIf(((SecTick Mod 3600) Mod 60) >= 10, ((SecTick Mod 3600) Mod 60) & "s", "0" & (SecTick Mod 3600) Mod 60 & "s")
            End If
        Else
            Exit Do
        End If
        
        DoEvents
        If blnDownLoad = False Then
            ' 下载被中止
            bStop
            DownloadToFile = False
            RaiseEvent Error(e4)
            RaiseEvent DownloadEnd(False)
            Exit Do
        End If
    Loop
    '关闭文件
    CloseHandle FileHandle
    
    If blnDownLoad = True Then
        '下载完成
        bStop
        DownloadToFile = True
        RaiseEvent Status(s6)
        RaiseEvent DownloadEnd(True)
        frmDownload.blnStop.Enabled = False
        frmDownload.blnStart.Enabled = True
    End If
End Function
