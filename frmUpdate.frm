VERSION 5.00
Begin VB.Form frmUpdate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "AG Updater"
   ClientHeight    =   2130
   ClientLeft      =   2250
   ClientTop       =   1335
   ClientWidth     =   5025
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
   ScaleHeight     =   2130
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLink 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      MouseIcon       =   "frmUpdate.frx":5F32
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label lblNewVer 
      Alignment       =   2  'Center
      Caption         =   "Checking For Updates.."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim InetFail As Boolean
'Const Delim As String = "*|\/<>"
'Const FileDelim As String = "|*\/|"
'Dim PatcherArg As String
Dim RunNum As Integer
'Dim FileClose As Boolean
'Dim FileStillExec As Boolean
'Dim Fso As New FileSystemObject
Dim strHomepage As String


Private WithEvents wd As WininetDown
Attribute wd.VB_VarHelpID = -1


'Private Sub lstAdd(ByVal lstin As String)
'lstStat.AddItem lstin
'lstStat.ListIndex = lstStat.NewIndex
'End Sub

Private Sub Form_Activate()
    Dim rectForm As RECT, objForm As Object
    SystemParametersInfo SPI_GETWORKAREA, 0, rectForm, 0
    Me.Move Screen.Width - Me.Width, rectForm.Bottom * Screen.TwipsPerPixelX - Me.Height
    RoundRect Me.hDC, 0, 0, Me.Width / Screen.TwipsPerPixelX - 1, Me.ScaleHeight / Screen.TwipsPerPixelY - 1, 10, 10
    For Each objForm In Me
        DoEvents
    Next
    On Error Resume Next
    
    If RunNum = 0 Then
        RunNum = 1
    Else
        Exit Sub
    End If
    
    CheckUpdate
    Exit Sub
    
    'lstAdd "Current Version: " & App.Major & "." & App.Minor & "." & App.Revision
    'lstAdd "Checking For Updates"
    'lstAdd "Downloading Config File"
    
    'Dim iniF() As Byte
    'iniF = InetUpdate.OpenURL("http://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini", icByteArray)
    'If InetFail = True Then Exit Sub
    
    'wd.URL = "dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini"
    'wd.FileName = Environ$("temp") & "\AGYTVGU.ini"
    'wd.DownloadToFile
    
    
    'lstAdd "Config File Downloaded"
    'lstAdd "Analyzing Config File"
    
    
    'SaveToFileFrm Environ$("temp") & "\AGYTVGU.ini", iniF
    'Open Environ$("temp") & "\AGYTVGU.ini" For Binary As #1
    'Dim X As Double
    'For x = 0 To UBound(iniF)
    'Put #1, , iniF(x)
    'DoEvents
    'Next
    'Close #1
    'Dim NewVer As String, tmpVer As String
    'NewVer = GetIni("Version", "Ver", App.Major & "." & App.Minor & "." & App.Revision, Environ$("temp") & "\AGYTVGU.ini")
    
    'tmpVer = Replace(NewVer, ".", "")
    'If Val(tmpVer) > (App.Major & App.Minor & App.Revision) Then
    'If Dir(App.Path & "\Backup\") = "" Then
    'MkDir App.Path & "\Backup"
    'Else
    ' Fso.DeleteFolder App.Path & "\Backup", True
    'End If
    'Dim OwnFileName As String
    'OwnFileName = App.EXEName
    
    'lstAdd "New Version Found: " & NewVer
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
    'lstAdd "Downloading Main File"
    'MainF = InetUpdate.OpenURL(GetIni("Patch", "Link", False, Environ$("temp") & "\AGYTVGU.ini"), icByteArray)
    'Dim MainFN As String
    'MainFN = GetIni("Patch", "Filename", False, Environ$("temp") & "\AGYTVGU.ini")
    'If InetFail = True Then Exit Sub
    'SaveToFileFrm Environ$("temp") & "\" & MainFN, MainF
    
    'wd.URL = GetIni("Patch", "Link", False, Environ$("temp") & "\AGYTVGU.ini")
    'wd.FileName = App.Path & "\" & NewVer & "-" & MainFN                    'Environ$("temp") & "\" & MainFN
    'If wd.DownloadToFile = True Then
    
    
    
    
    'If Dir(App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe") <> "" Then Kill App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe"
    'Name App.Path & "\" & App.EXEName & ".exe" As App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & App.EXEName & ".exe"
    'If Dir$(App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe") <> "" Then Kill App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe"
    'Name App.Path & "\" & OwnFileName & ".exe" As App.Path & "\Backup\" & App.Major & "." & App.Minor & "." & App.Revision & "-" & OwnFileName & ".exe"
    
    'Name App.Path & "\" & NewVer & "-" & MainFN As App.Path & "\" & OwnFileName & ".exe"
    'Open Environ$("temp") & "\" & MainFN For Binary As #3
    
    'For x = 0 To UBound(MainF)
    'Put #3, , MainF(x)
    'DoEvents
    'Next
    'Close #3
    
    'lstAdd "Main File Downloaded"
    'PatcherArg = PatcherArg & FileDelim & Environ$("temp") & "\" & MainFN
    'lstAdd "Checking For External Updates"
    
    'If Val(GetIni("OtherFile", "IsGotOtherFile", 0, Environ$("temp") & "\AGYTVGU.ini")) = 0 Then
    'lstAdd "No External Updates"
    'Else
    'Dim FileExDl() As String
    'lstAdd "External Updates Found"
    
    'Dim FileInd As Long
    'FileExDl = Split(GetIni("OtherFile", "FileDownloadLink", 0, Environ$("temp") & "\AGYTVGU.ini"), ",")
    
    'For X = LBound(FileExDl) To UBound(FileExDl)
    'Dim ExUpdateF() As Byte
    'InetUpdate.URL = FileExDl(FileInd)
    'MsgBox InetUpdate.GetHeader
    'lstAdd "Downloading " & FileExDl(FileInd + 1)
    'PatcherArg = PatcherArg & FileDelim & Environ$("temp") & "\" & FileExDl(FileInd + 1)
    
    'wd.URL = FileExDl(FileInd)
    'wd.FileName = App.Path & "\" & FileExDl(FileInd + 1)
    'If wd.DownloadToFile = True Then
    
    
    'ExUpdateF = InetUpdate.OpenURL(FileExDl(FileInd), icByteArray)
    'SaveToFileFrm Environ$("temp") & "\" & FileExDl(FileInd + 1), ExUpdateF
    'lstAdd FileExDl(FileInd + 1) & " Downloaded"
    'If (FileInd + 2) <= UBound(FileExDl) Then
    'FileInd = FileInd + 2
    'Else
    'Exit For
    'End If
    'DoEvents
    'End If
    'Next
    
    
    'End If
    'lstAdd "Done Downloading Updates"
    'lstAdd "Program Restarting..."
    
    'Do Until OpenProcess(&H400, 0, Shell(App.Path & "\" & OwnFileName & ".exe")) <> 0
    'DoEvents
    'Loop
    'Shell "cmd /c start """ & App.Path & "\" & OwnFileName & ".exe" & """", vbHide
    'End
    'lstAdd "Patching Updates"
    'PatcherArg = PatcherArg & Delim & App.Path & "\" & Delim & NewVer & Delim & App.Major & "." & App.Minor & "." & App.Revision
    'If Me.Visible = True Then
    'Shell Environ$("temp") & "\AG.exe " & PatcherArg, vbNormalFocus
    'End
    'End If
    
    'End If
    'Else
    
    'lstAdd "The Current Version Is The Latest Version"
    'frmMain.SetFocus
    'Dim nTime As Long
    'nTime = Timer()
    'Do
    'DoEvents
    'If Timer() - nTime > 3 Then Unload Me
    'Loop
    
    'End If
End Sub

Private Sub Form_Load()
    Set wd = New WininetDown
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 10, 10), True
    
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

'Private Sub wd_Progress(ByVal dbFileSize As Double, ByVal dbFinished As Double, ByVal dbSpeed As Double)
'Dim BFile       As String
'Dim BRead       As String
'Dim BShare      As Double
'Dim BSpeed      As String

'If IsNumeric(dbFinished) = True And IsNumeric(dbFileSize) = True Then
'BFile = VBStrFormatByteSize(dbFileSize)
'BRead = VBStrFormatByteSize(dbFinished)
'BShare = dbFinished / (dbFileSize / 100)
'If BShare < 0 Then BShare = 0
'If BShare > 100 Then BShare = 100
'lblProgress.Caption = Round(BShare, 2) & "%"
'myProgressBar1.Value = BShare
'Downloading.Caption = BRead & " of " & BFile & " transferred (" & Round(BShare, 2) & "%)"
'End If
'If IsNumeric(dbSpeed) = True Then
'BSpeed = VBStrFormatByteSize(dbSpeed)
'DownSpeed.Caption = "Throughput " & BSpeed & " /s"
'End If
'End Sub

'Private Sub wd_Status(ByVal lStatus As wd_Status)
'Dim strMsg As String
'Select Case lStatus
'Case 1
'strMsg = "Starting connection..."
'Case 2
'strMsg = "Connecting to server..."
'Case 3
'strMsg = "Sending request..."
'Case 4
'strMsg = "Done sending request..."
'Case 5
' strMsg = "Receiving data..."
'Case 6
'strMsg = "Done Downloading..."
'Case Else
'strMsg = "Unknow status..."
'End Select
'If strMsg <> lstStat.List(lstStat.ListCount - 1) Then
'lstAdd strMsg
'End If
'End Sub


'Private Sub wd_Error(ByVal lErrorCode As wd_ErrorCode)
'Dim strMsg As String
'Select Case lErrorCode
'Case 1
'strMsg = "Connection failed..."
'Case 2
'strMsg = "Failed to connect to the server..."
'Case 3
'strMsg = "Failed to send request..."
'Case 4
'strMsg = "Download aborted..."
'Case 5
'strMsg = "File exist..."
'Case 6
'strMsg = "Failed to create directory..."
'Case Else
'strMsg = "Failed to open file..."
'End Select

'If strMsg <> lstStat.List(lstStat.ListCount - 1) Then
'lstAdd strMsg
'End If
'End Sub


Private Function CheckUpdate() As Boolean
    wd.URL = "dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini"
    wd.FileName = Environ$("temp") & "\AGYTVGU.ini"
    wd.DownloadToFile
    Dim strNewVer() As String, strNewVersion As String, strCurrentVersion As String, nTime As Long
    strCurrentVersion = App.Major & "." & App.Minor & "." & App.Revision
    strNewVersion = GetIni("Version", "Ver", strCurrentVersion, Environ$("temp") & "\AGYTVGU.ini")
    strNewVer = Split(strNewVersion, ".")
    If Val(strNewVer(0)) < App.Major Then
        GoTo NoUpdates
    ElseIf Val(strNewVer(0)) = App.Major Then
        If Val(strNewVer(1)) < App.Minor Then
            GoTo NoUpdates
        ElseIf Val(strNewVer(1)) = App.Minor Then
            If Val(strNewVer(2)) < App.Revision Then
                GoTo NoUpdates
            End If
        End If
    End If
    If Me.Visible = False Then Me.Show
    lblNewVer.Caption = "Found new version: " & strNewVersion & vbCrLf & "  Current Version: " & strCurrentVersion
    strHomepage = GetIni("Updatelink", "UpdateHomePage", "http://garyngzhongbo.blogspot.com", Environ$("temp") & "\AGYTVGU.ini")
    lblLink.Caption = "Check and download the new version here: " & vbCrLf & strHomepage
    nTime = Timer()
    Do Until Timer() - nTime > 5
        DoEvents
    Loop
    VisitURL strHomepage
    Unload Me
    Exit Function
NoUpdates:
    CheckUpdate = False
    lblNewVer.Caption = "The current version " & App.Major & "." & App.Minor & "." & App.Revision & "is the latest verison."
    Do Until Timer() - nTime > 5
        DoEvents
    Loop
    Unload Me
End Function


Private Sub lblLink_Click()
    VisitURL strHomepage
    Unload Me
End Sub
