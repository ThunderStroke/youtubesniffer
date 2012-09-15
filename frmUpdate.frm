VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG Updater"
   ClientHeight    =   3630
   ClientLeft      =   2295
   ClientTop       =   1770
   ClientWidth     =   6135
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AGYouTubeVideoGrabber.SFProgressBar myProgressBar1 
      Height          =   150
      Left            =   120
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   265
   End
   Begin VB.ListBox lstStat 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Downloading 
      Caption         =   "0 of 0 bytes transferred (0%)"
      Height          =   225
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label LbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Download Progress:"
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label LbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Download Speed:"
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label DownSpeed 
      Caption         =   "Throughput 0.00 KB/s"
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InetFail As Boolean
Const Delim As String = "*|\/<>"
Const FileDelim As String = "|*\/|"
'Dim PatcherArg As String
Dim RunNum As Integer
'Dim FileClose As Boolean
'Dim FileStillExec As Boolean
'Dim Fso As New FileSystemObject



Private WithEvents wd As WininetDown
Attribute wd.VB_VarHelpID = -1


Private Sub lstAdd(ByVal lstin As String)
    lstStat.AddItem lstin
    lstStat.ListIndex = lstStat.NewIndex
End Sub

Private Sub cmdCancel_Click()

End Sub



Private Sub Form_Activate()
    'Exit Sub
    On Error Resume Next
    
    If RunNum = 0 Then
        RunNum = 1
    Else
        Exit Sub
    End If
    
    lstAdd "Current Version: " & App.Major & "." & App.Minor & "." & App.Revision
    lstAdd "Checking For Updates"
    lstAdd "Downloading Config File"
    
    'Dim iniF() As Byte
    'iniF = InetUpdate.OpenURL("http://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini", icByteArray)
    'If InetFail = True Then Exit Sub
    
    wd.URL = "dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini"
    wd.FileName = Environ$("temp") & "\AGYTVGU.ini"
    wd.DownloadToFile
    
    
    lstAdd "Config File Downloaded"
    lstAdd "Analyzing Config File"
    
    
    'SaveToFileFrm Environ$("temp") & "\AGYTVGU.ini", iniF
    'Open Environ$("temp") & "\AGYTVGU.ini" For Binary As #1
    Dim x As Double
    'For x = 0 To UBound(iniF)
    'Put #1, , iniF(x)
    'DoEvents
    'Next
    'Close #1
    Dim NewVer As String, tmpVer As String
    NewVer = GetIni("Version", "Ver", App.Major & "." & App.Minor & "." & App.Revision, Environ$("temp") & "\AGYTVGU.ini")
    
    tmpVer = Replace(NewVer, ".", "")
    If Val(tmpVer) > (App.Major & App.Minor & App.Revision) Then
        If Dir(App.Path & "\Backup\") = "" Then
            MkDir App.Path & "\Backup"
            'Else
            ' Fso.DeleteFolder App.Path & "\Backup", True
        End If
        Dim OwnFileName As String
        OwnFileName = App.EXEName
        
        lstAdd "New Version Found: " & NewVer
        'If Val(GetIni("BigUpdate", "ManualDownload", 0, Environ$("temp") & "\AGYTVGU.ini")) <> 0 Then
        'lstAdd "Current Updates Is A Big Update Please Download Manually"
        'End If
        
        'lstAdd "Downloading Patcher"
        
        'PatchF = InetUpdate.OpenURL(GetIni("Patcher", "PatcherLink", "http://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/AG%20Patcher.exe", Environ$("temp") & "\AGYTVGU.ini"), icByteArray)
        'If InetFail = True Then Exit Sub
        'SaveToFileFrm Environ$("temp") & "\AG.exe", PatchF
        'wd.URL = GetIni("Patcher", "PatcherLink", "http://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/AG%20Patcher.exe", Environ$("temp") & "\AGYTVGU.ini")
        'wd.FileName = Environ$("temp") & "\AG.exe"
        'wd.DownloadToFile
        
        
        'Open Environ$("temp") & "\AG Patcher.exe" For Binary As #2
        
        'For x = 0 To UBound(PatchF)
        'Put #2, , PatchF(x)
        'DoEvents
        'Next
        'Close #2
        'lstAdd "Patcher Downloaded"
        lstAdd "Downloading Main File"
        'MainF = InetUpdate.OpenURL(GetIni("Patch", "Link", False, Environ$("temp") & "\AGYTVGU.ini"), icByteArray)
        Dim MainFN As String
        MainFN = GetIni("Patch", "Filename", False, Environ$("temp") & "\AGYTVGU.ini")
        'If InetFail = True Then Exit Sub
        'SaveToFileFrm Environ$("temp") & "\" & MainFN, MainF
        
        wd.URL = GetIni("Patch", "Link", False, Environ$("temp") & "\AGYTVGU.ini")
        wd.FileName = App.Path & "\" & NewVer & "-" & MainFN                    'Environ$("temp") & "\" & MainFN
        If wd.DownloadToFile = True Then
            
            
            
            
            'If Dir(App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe") <> "" Then Kill App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe"
            'Name App.Path & "\" & App.EXEName & ".exe" As App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe"
            If Dir$(App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe") <> "" Then Kill App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe"
            Name App.Path & "\" & OwnFileName & ".exe" As App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe"
            
            Name App.Path & "\" & NewVer & "-" & MainFN As App.Path & "\" & OwnFileName & ".exe"
            'Open Environ$("temp") & "\" & MainFN For Binary As #3
            
            'For x = 0 To UBound(MainF)
            'Put #3, , MainF(x)
            'DoEvents
            'Next
            'Close #3
            
            lstAdd "Main File Downloaded"
            'PatcherArg = PatcherArg & FileDelim & Environ$("temp") & "\" & MainFN
            lstAdd "Checking For External Updates"
            
            If Val(GetIni("OtherFile", "IsGotOtherFile", 0, Environ$("temp") & "\AGYTVGU.ini")) = 0 Then
                lstAdd "No External Updates"
            Else
                Dim FileExDl() As String
                lstAdd "External Updates Found"
                
                Dim FileInd As Long
                FileExDl = Split(GetIni("OtherFile", "FileDownloadLink", 0, Environ$("temp") & "\AGYTVGU.ini"), ",")
                
                For x = LBound(FileExDl) To UBound(FileExDl)
                    'Dim ExUpdateF() As Byte
                    'InetUpdate.URL = FileExDl(FileInd)
                    'MsgBox InetUpdate.GetHeader
                    lstAdd "Downloading " & FileExDl(FileInd + 1)
                    'PatcherArg = PatcherArg & FileDelim & Environ$("temp") & "\" & FileExDl(FileInd + 1)
                    
                    wd.URL = FileExDl(FileInd)
                    wd.FileName = App.Path & "\" & FileExDl(FileInd + 1)
                    If wd.DownloadToFile = True Then
                        
                        
                        'ExUpdateF = InetUpdate.OpenURL(FileExDl(FileInd), icByteArray)
                        'SaveToFileFrm Environ$("temp") & "\" & FileExDl(FileInd + 1), ExUpdateF
                        lstAdd FileExDl(FileInd + 1) & " Downloaded"
                        If (FileInd + 2) <= UBound(FileExDl) Then
                            FileInd = FileInd + 2
                        Else
                            Exit For
                        End If
                        DoEvents
                    End If
                Next
                
                
            End If
            lstAdd "Done Downloading Updates"
            lstAdd "Program Restarting..."
            
            Do Until OpenProcess(&H400, 0, Shell(App.Path & "\" & OwnFileName & ".exe")) <> 0
                DoEvents
            Loop
            'Shell "cmd /c start """ & App.Path & "\" & OwnFileName & ".exe" & """", vbHide
            End
            'lstAdd "Patching Updates"
            'PatcherArg = PatcherArg & Delim & App.Path & "\" & Delim & NewVer & Delim & App.Major & "." & App.Minor & "." & App.Revision
            'If Me.Visible = True Then
            'Shell Environ$("temp") & "\AG.exe " & PatcherArg, vbNormalFocus
            'End
            'End If
            
        End If
    Else
        
        lstAdd "The Current Version Is The Latest Version"
        Dim t As Long
        t = Timer()
        Do
            DoEvents
            If Timer() - t > 3 Then Unload Me
        Loop
        
    End If
End Sub

Private Sub Form_Load()
    Set wd = New WininetDown
    Attach Me.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If FileStillExec = True Then
    'FileClose = True
    'Cancel = 1
    'MsgBox "File Still Executing!" & vbCrLf & "Close after a few seconds.", vbExclamation
    'Else
    'InetUpdate.Cancel
    'Cancel = 0
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RunNum = 0
    Detach Me.hWnd
    Unload Me
End Sub


'Private Sub InetUpdate_StateChanged(ByVal State As Integer)
'On Error GoTo goterr
'If State = icError Then
'lstAdd "Download File Failed"
'InetFail = True
'End If
'Exit Sub
'goterr:
'InetUpdate.Cancel
'InetFail = True
'Exit Sub

'End Sub
'Private Sub SaveToFileFrm(ByVal DirFilename As String, ByRef FileByte() As Byte)
'Open DirFilename For Binary As #3
'Dim x As Double
'FileStillExec = True
'For x = 0 To UBound(FileByte)
'If FileClose = True Then
'Close #3
'Unload Me
'End If
'Put #3, , FileByte(x)
'DoEvents
'Next
'Close #3
'FileStillExec = False
'End Sub






' 下载进度
Private Sub wd_Progress(ByVal dbFileSize As Double, ByVal dbFinished As Double, ByVal dbSpeed As Double)
    Dim BFile       As String
    Dim BRead       As String
    Dim BShare      As Double
    Dim BSpeed      As String
    
    If IsNumeric(dbFinished) = True And IsNumeric(dbFileSize) = True Then
        BFile = VBStrFormatByteSize(dbFileSize)
        BRead = VBStrFormatByteSize(dbFinished)
        BShare = dbFinished / (dbFileSize / 100)
        If BShare < 0 Then BShare = 0
        If BShare > 100 Then BShare = 100
        myProgressBar1.Value = BShare
        Downloading.Caption = BRead & " of " & BFile & " transferred (" & Round(BShare, 2) & "%)"
    End If
    If IsNumeric(dbSpeed) = True Then
        BSpeed = VBStrFormatByteSize(dbSpeed)
        DownSpeed.Caption = "Throughput " & BSpeed & " /s"
    End If
End Sub


' 下载结束
Private Sub wd_DownloadEnd(ByVal bSuccess As Boolean)
    'Debug.Print bSuccess
End Sub

' 下载状态
Private Sub wd_Status(ByVal lStatus As wd_Status)
    Dim strMsg As String
    Select Case lStatus
    Case 1
        strMsg = "Starting connection..."
    Case 2
        strMsg = "Connecting to server..."
    Case 3
        strMsg = "Sending request..."
    Case 4
        strMsg = "Done sending request..."
    Case 5
        strMsg = "Receiving data..."
    Case 6
        strMsg = "Done Downloading..."
    Case Else
        strMsg = "Unknow status..."
    End Select
    If strMsg <> lstStat.List(lstStat.ListCount - 1) Then
        lstAdd strMsg
    End If
End Sub

' 下载出错
Private Sub wd_Error(ByVal lErrorCode As wd_ErrorCode)
    Dim strMsg As String
    Select Case lErrorCode
    Case 1
        strMsg = "Connection failed..."
    Case 2
        strMsg = "Failed to connect to the server..."
    Case 3
        strMsg = "Failed to send request..."
    Case 4
        strMsg = "Download aborted..."
    Case 5
        strMsg = "File exist..."
    Case 6
        strMsg = "Failed to create directory..."
    Case Else
        strMsg = "Failed to open file..."
    End Select
    
    If strMsg <> lstStat.List(lstStat.ListCount - 1) Then
        lstAdd strMsg
    End If
End Sub


