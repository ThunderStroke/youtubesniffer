VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG Video Converter"
   ClientHeight    =   4770
   ClientLeft      =   7395
   ClientTop       =   5115
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProgress 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2760
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Video Quality"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   5055
      Begin VB.ComboBox cmbQuality 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmConvert.frx":5F32
         Left            =   240
         List            =   "frmConvert.frx":5F42
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Video Format"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   5055
      Begin VB.ComboBox cmbFormat 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmConvert.frx":5F62
         Left            =   240
         List            =   "frmConvert.frx":5F64
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Start Convert!"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cdInOut 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtOut 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox txtIn 
      Height          =   315
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Output Video Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Input Video Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuExpertMode 
         Caption         =   "Expert Mode"
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open Folder After Converting"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Save Log File"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetShortPathName Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Dim ConverterFileEx As String
Dim ConvertQuality As String
Public ConvertPara As String
'Dim FileShortP As String
Dim lThreadHandle1 As Long
Dim lpThreadID As Long
Public TimeTick As Long


Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_TERMINATE = (&H1)


Private Sub cmbFormat_Click()
    Select Case cmbFormat.ListIndex                                             'cmbFormat.List(cmbFormat.ListIndex)
    Case 0
        ConverterFileEx = ".mov"
    Case 1
        ConverterFileEx = ".mp4"
    Case 2
        ConverterFileEx = ".mp4"
    Case 3
        ConverterFileEx = ".3gp"
    Case 4
        ConverterFileEx = ".wmv"
    Case 5
        ConverterFileEx = ".avi"
    Case 6
        ConverterFileEx = ".mp3"
    End Select
    txtOut.Text = txtIn.Text & ConverterFileEx
    cmdSave.Enabled = True
    If cmbQuality.List(cmbQuality.ListIndex) <> "" Then
        cmdConvert.Enabled = True
    End If
End Sub

Private Sub cmbQuality_Click()
    Select Case cmbQuality.ListIndex
        
    Case 0
        ConvertQuality = " -qscale 1 "
    Case 1
        ConvertQuality = " -qscale 7 "
    Case 2
        ConvertQuality = " -qscale 15 "
    Case 3
        ConvertQuality = " "
    End Select
    
    If cmbFormat.List(cmbFormat.ListIndex) <> "" Then
        cmdConvert.Enabled = True
    End If
End Sub



Private Sub cmdConvert_Click()
    cmdStop.Enabled = True
    TimeTick = Timer()
    
    
    txtProgress.Text = "Conversion started at: " & Now
    txtProgress.SelStart = Len(txtProgress.Text)
    Select Case cmbFormat.ListIndex
    Case 0
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -r 24.00 -b 500.00k " & ConvertQuality & "-ar 22050 -ab 64.00k -ac 2 -vol 655  -y " & """" & txtOut.Text & """"
    Case 1
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -vcodec mpeg4 -r 29.00 -b 1500.00k " & ConvertQuality & "-ar 44100 -ab 128.00k -ac 2 -vol 655  -y " & """" & txtOut.Text & """"
    Case 2
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -f psp -vcodec mpeg4 -r 29.00 -b 500.00k " & ConvertQuality & "-ar 24000 -ab 64.00k -ac 2 -vol 655  -y " & """" & txtOut.Text & """"
    Case 3
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -vcodec mpeg4 -r 15.00 -b 30.00k " & ConvertQuality & "-ar 8000 -ab 12.20k -ac 1 -vol 655  -y " & """" & txtOut.Text & """"
    Case 4
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -f asf -vcodec wmv2 -r 15.00 -b 50.00k " & ConvertQuality & "-acodec wmav2 -ar 22050 -ab 64.00k -ac 2 -vol 655  -y " & """" & txtOut.Text & """"
    Case 5
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -vcodec mpeg4 -vtag XVID -b 200.00k " & ConvertQuality & "-ar 22050 -ab 64.00k -ac 2 -vol 655  -y  " & """" & txtOut.Text & """"
    Case 6
        ConvertPara = "ffmpeg -i " & """" & txtIn.Text & """" & " -vn -ar 48000 -ab 256.00k -ac 2 -vol 655  -y  " & """" & txtOut.Text & """"
    End Select
    cmdConvert.Enabled = False
    cmdIn.Enabled = False
    cmdSave.Enabled = False
    
    cmbFormat.Enabled = False
    cmbQuality.Enabled = False
    'Debug.Print ConvertPara
    If mnuExpertMode.Checked = False Then
        'pgBar.Value = 10
        WaitExeRun ConvertPara
        'pgBar.Value = 100
        txtProgress.Text = txtProgress.Text & "Done converting at: " & Now & ". Time taken: " & Format(Timer() - frmConvert.TimeTick, "0.##") & " seconds"
        txtProgress.SelStart = Len(txtProgress.Text)
        
        cmdConvert.Enabled = True
        cmbFormat.Enabled = True
        cmbQuality.Enabled = True
        cmdConvert.Enabled = True
        cmdIn.Enabled = True
        cmdSave.Enabled = True
        cmdStop.Enabled = False
        If mnuOpenFolder.Checked = True Then
            Shell "explorer /select," & txtOut.Text, vbNormalFocus
        End If
        
    Else
        'pgBar.Value = 10
        lThreadHandle1 = CreateThread(ByVal 0&, ByVal 0&, AddressOf ExecuteApp, ByVal 0&, 0, lpThreadID)
        lEventHandle = CreateEvent(ByVal 0&, False, False, ByVal 0&)
        
    End If
    
    
End Sub

Private Sub cmdIn_Click()
    'Dim FileBuffer As String
    
    With cdInOut
        .ShowOpen
        .Filter = "*.*|*.*"
        If Len(.FileName) = 0 Then
            Exit Sub
        Else
            'FileBuffer = Space(Len(.Filename))
            txtIn.Text = .FileName                                              'GetShortPathName(.Filename, FileShortP, Len(.Filename) + 1)
        End If
    End With
    
End Sub

Private Sub cmdSave_Click()
    With cdInOut
        .Filter = cmbFormat.List(cmbFormat.ListIndex) & "|" & "*" & ConverterFileEx
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        Else
            txtOut.Text = .FileName
        End If
    End With
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    If mnuExpertMode.Checked = True Then
        Call TerminateThread(lThreadHandle1, ByVal 0&)
        TerminateProcess OpenProcess(1&, True, proc.dwProcessId), 0&
    Else
        TerminateProcess WaithProcess, 0&
    End If
    txtProgress.Text = txtProgress.Text & vbCrLf & "Conversion aborted by user."
    txtProgress.SelStart = Len(txtProgress.Text)
    cmdConvert.Enabled = True
    cmbFormat.Enabled = True
    cmbQuality.Enabled = True
    cmdConvert.Enabled = True
    cmdIn.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub Form_Load()
    
    If App.LogMode = 0 Then
        MsgBox "Don't Run At IDE Please!", vbInformation
        Unload Me
        Exit Sub
    End If
    
    Dim x As Integer, ListArr(6) As String
    ListArr(0) = "Ipod VIdeo (Apple Quick Time *.MOV)"
    ListArr(1) = "Iphone Video (MPEG-4 *.MP4)"
    ListArr(2) = "PSP Video (H.264 *.MP4)"
    ListArr(3) = "Cell Phone (H.264 *.3GP)"
    ListArr(4) = "Windows Media Video (v.7 *.WMV)"
    ListArr(5) = "XVid MPEG-4 Codec (*.AVI)"
    ListArr(6) = "MPEG Audio Layer 3 (*.MP3)"
    
    For x = LBound(ListArr) To UBound(ListArr)
        cmbFormat.AddItem ListArr(x)
    Next
    'mnuLog.Checked = GetIni("Converter", "SaveLogFile", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuOpenFolder.Checked = GetIni("Converter", "OpenFolder", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuExpertMode.Checked = GetIni("Converter", "ExpertMode", True, App.Path & "\YoutubeGrabberOption.ini")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If frmMain.Visible = False Then
        
        'Call TerminateProcess(GetCurrentProcess, ByVal 0&)
        End
    End If
    If lThreadHandle1 > 0 Then
        
        Call TerminateThread(lThreadHandle1, ByVal 0&)
        TerminateProcess OpenProcess(1&, True, proc.dwProcessId), 0&
    End If
End Sub



Private Sub mnuExpertMode_Click()
    If mnuExpertMode.Checked = True Then
        mnuExpertMode.Checked = False
        WriteKey "Converter", "ExpertMode", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuExpertMode.Checked = True
        WriteKey "Converter", "ExpertMode", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

'Private Sub mnuLog_Click()
'If mnuLog.Checked = True Then
'mnuLog.Checked = False
'WriteKey "Converter", "SaveLogFile", False, App.Path & "\YoutubeGrabberOption.ini"
'Else
'mnuLog.Checked = True
'WriteKey "Converter", "SaveLogFile", True, App.Path & "\YoutubeGrabberOption.ini"
'End If
'End Sub

Private Sub mnuOpenFolder_Click()
    If mnuOpenFolder.Checked = True Then
        mnuOpenFolder.Checked = False
        WriteKey "Converter", "OpenFolder", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuOpenFolder.Checked = True
        WriteKey "Converter", "OpenFolder", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub



Private Sub txtIn_Change()
    If Len(txtIn.Text) <> 0 Then
        cmbFormat.Enabled = True
        cmbQuality.Enabled = True
    Else
        cmbFormat.Enabled = False
        cmbQuality.Enabled = False
    End If
    
    If cmbFormat.List(cmbFormat.ListIndex) <> "" And cmbQuality.List(cmbQuality.ListIndex) <> "" Then
        cmbFormat_Click
    End If
End Sub

Private Sub txtOut_Change()
    If Dir$(txtOut.Text) <> "" Then
        Static Num As Integer
        Dim TextTmp As String
        If TextTmp = "" Then
            TextTmp = txtOut.Text
        End If
        Do Until Dir$(txtOut.Text) = ""
            If Num = 0 Then Num = 1
            txtOut.Text = Mid$(TextTmp, 1, InStrRev(TextTmp, ".") - 1) & "(" & Num & ")" & Right$(TextTmp, Len(TextTmp) - InStrRev(TextTmp, ".") + 1)
            Num = Num + 1
        Loop
    End If
    
End Sub
