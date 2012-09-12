VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlay 
   Caption         =   "AG Video Player"
   ClientHeight    =   5025
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "While Playing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   3255
      Begin VB.Label lblHelp 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      TabIndex        =   10
      Top             =   1440
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   6600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Play Network Video/Audio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   6615
      Begin VB.CommandButton cmdAddNetworkLink 
         Caption         =   " Add To Play List"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txtLink 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Link:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drag Files To The Playlist "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdSelectFile 
         Caption         =   "..."
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtFileDir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2520
         Width           =   5655
      End
      Begin MSComctlLib.ListView lvwPlayList 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Format"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPlaySelected 
      Caption         =   "Play Selected Item"
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play All"
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSFile 
         Caption         =   "Check For Supported File Type"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuFFOpt 
         Caption         =   "ffplay Option"
         Begin VB.Menu mnuNoVideo 
            Caption         =   "Disable Video"
         End
         Begin VB.Menu mnuNoAudio 
            Caption         =   "Disable Audio "
         End
         Begin VB.Menu mnuFullScreen 
            Caption         =   "Full Screen"
         End
         Begin VB.Menu mnuShowMode 
            Caption         =   "Show Mode"
            Begin VB.Menu mnuSMVideo 
               Caption         =   "Video"
            End
            Begin VB.Menu mnuSMWaves 
               Caption         =   "Waves"
            End
            Begin VB.Menu mnuSMRDFT 
               Caption         =   "RDFT"
            End
         End
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPlaySelected 
         Caption         =   "Play Selected Item"
      End
      Begin VB.Menu mnuDelSelected 
         Caption         =   "Delete Selected Item"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FsoSys As New FileSystemObject
Dim codecs As String



Private Sub cmdAddNetworkLink_Click()
    Dim lx As ListItem
    Set lx = lvwPlayList.ListItems.Add(, , txtLink.Text)
    lx.SubItems(2) = "Network File"
    
End Sub

Private Sub cmdExit_Click()
    TerminateProcess OpenProcess(1&, True, ffplayPID), 0&
    lvwPlayList.SelectedItem.SubItems(1) = ""
End Sub

Private Sub cmdPlay_Click()
    Dim x As Long
    For x = 0 To lvwPlayList.ListItems.Count - 1
        lvwPlayList.SetFocus
        lvwPlayList.SelectedItem.Selected = x
        Play
        DoEvents
    Next
End Sub

Private Sub cmdPlaySelected_Click()
    Play
End Sub

Private Sub cmdSelectFile_Click()
    Dim lx As ListItem
    With cdFile
        .DialogTitle = "Select File To Add To Play List..."
        .Filter = "All File Types(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        If Dir$(.FileName) = "" Then Exit Sub
        Set lx = lvwPlayList.ListItems.Add(, , .FileName)
        txtFileDir.Text = .FileName
        
        If SupportOrNot(FsoSys.GetExtensionName(.FileName)) Then
            lx.SubItems(2) = FsoSys.GetExtensionName(.FileName)
        Else
            lx.SubItems(2) = "Not Supported"
        End If
        
        
    End With
End Sub


Private Sub Form_Load()
    codecs = "|/\<>*.3g2|/\<>*.3gp|/\<>*.4xm|/\<>*.IFF|/\<>*.ISS|/\<>*.MTV|/\<>*.RoQ|/\<>*.a64|/\<>*.aac|/\<>*.ac3|/\<>*.act|/\<>*.adf|/\<>" & _
    "*.adts|/\<>*.adx|/\<>*.aea|/\<>*.aiff|/\<>*.alaw|/\<>*.amr|/\<>*.anm|/\<>*.apc|/\<>*.ape|/\<>*.asf|/\<>*.asf_stream|/\<>" & _
    "*.ass|/\<>*.au|/\<>*.avi|/\<>*.avm2|/\<>*.avs|/\<>*.bethsoftvid|/\<>*.bfi|/\<>*.bin|/\<>*.bink|/\<>*.bit|/\<>*.bmv|/\<>*.c93|/\<>" & _
    "*.caf|/\<>*.cavsvideo|/\<>*.cdg|/\<>*.cdxl|/\<>*.crc|/\<>*.daud|/\<>*.dfa|/\<>*.dirac|/\<>*.dnxhd|/\<>*.dshow|/\<>*.dsicin|/\<>*.dts|/\<>*.dv|/\<>*.dvd|/\<>*.dxa|/\<>*.ea|/\<>" & _
    "*.ea_cdata|/\<>*.eac3|/\<>*.f32be|/\<>*.f32le|/\<>*.f64be|/\<>*.f64le|/\<>*.ffm|/\<>*.ffmetadata|/\<>*.film_cpk|/\<>*.filmstrip|/\<>*.flac|/\<>*.flic|/\<>*.flv|/\<>*.framecrc|/\<>*.framemd5|/\<>*.g722|/\<>*.g723_1|/\<>*.g729|/\<>*.gif|/\<>*.gsm|/\<>*.gxf|/\<>*.h261|/\<>*.h263|/\<>*.h264|/\<>*.hls,applehttp|/\<>*.ico|/\<>*.idcin|/\<>*.idf|/\<>*.image2|/\<>*.image2pipe|/\<>*.ingenient|/\<>*.ipmovie|/\<>*.ipod|/\<>*.ismv|/\<>*.iv8|/\<>*.ivf|/\<>" & _
    "*.jv|/\<>*.latm|/\<>*.lavfi|/\<>*.lmlm4|/\<>*.loas|/\<>*.lxf|/\<>*.m4v|/\<>*.matroska|/\<>*.matroska,webm|/\<>*.md5|/\<>*.microdvd|/\<>*.mjpeg|/\<>*.mkvtimestamp_v2|/\<>*.mlp|/\<>*.mm|/\<>*.mmf|/\<>*.mov|/\<>*.mov,mp4,m4a,3gp,3g2,mj2|/\<>*.mp2|/\<>*.mp3|/\<>*.mp4|/\<>*.mpc|/\<>*.mpc8|/\<>*.mpeg|/\<>*.mpeg1video|/\<>*.mpeg2video|/\<>*.mpegts|/\<>*.mpegtsraw|/\<>*.mpegvideo|/\<>*.mpjpeg|/\<>*.msnwctcp|/\<>*.mulaw|/\<>*.mvi|/\<>*.mxf|/\<>*.mxf_d10|/\<>*.mxg|/\<>*.nc|/\<>*.nsv|/\<>*.null|/\<>*.nut|/\<>*.nuv|/\<>*.ogg|/\<>*.oma|/\<>*.pmp|/\<>*.psp|/\<>*.psxstr|/\<>*.pva|/\<>*.qcp|/\<>*.r3d|/\<>*.rawvideo|/\<>*.rcv|/\<>*.rl2|/\<>*.rm|/\<>*.rpl|/\<>*.rso|/\<>*.rtp|/\<>*.rtsp|/\<>*.s16be|/\<>*.s16le|/\<>*.s24be|/\<>*.s24le|/\<>*.s32be|/\<>*.s32le|/\<>*.s8|/\<>*.sap|/\<>*.sbg|/\<>*.sdl|/\<>*.sdp|/\<>*.segment|/\<>*.shn|/\<>*.siff|/\<>*" & _
    ".smjpeg|/\<>*.smk|/\<>*.sol|/\<>*.sox|/\<>*.spdif|/\<>*.srt|/\<>*.svcd|/\<>*.swf|/\<>*.thp|/\<>*.tiertexseq|/\<>*.tmv|/\<>*.truehd|/\<>*.tta|/\<>*.tty|/\<>*.txd|/\<>*.u16be|/\<>*.u16le|/\<>*.u24be|/\<>*.u24le|/\<>*.u32be|/\<>*.u32le|/\<>*.u8|/\<>*.vc1|/\<>*.vc1test|/\<>*.vcd|/\<>*.vfwcap|/\<>*.vmd|/\<>*.vob|/\<>*.voc|/\<>*.vqf|/\<>*.w64|/\<>*.wav|/\<>*.wc3movie|/\<>*.webm|/\<>*.wsaud|/\<>*.wsvqa|/\<>*.wtv|/\<>*.wv|/\<>*.xa|/\<>*.xbin|/\<>*.xmv|/\<>*.xwma|/\<>*.yop|/\<>*.yuv4mpegpipe"
    
    
    Select Case GetIni("Player", "ShowMode", 0, App.Path & "\YoutubeGrabberOption.ini")
    Case 0
        mnuSMVideo.Checked = True
    Case 1
        mnuSMWaves.Checked = True
    Case 2
        mnuSMRDFT.Checked = True
    End Select
    
    mnuFullScreen.Checked = GetIni("Player", "FullScreen", False, App.Path & "\YoutubeGrabberOption.ini")
    mnuNoVideo.Checked = GetIni("Player", "NoVideo", False, App.Path & "\YoutubeGrabberOption.ini")
    mnuNoAudio.Checked = GetIni("Player", "NoAudio", False, App.Path & "\YoutubeGrabberOption.ini")
    
    lblHelp.Caption = "p/space - Pause/Play" & vbCrLf & "f - Toggle Full Screen" & vbCrLf & "q - Quit" & vbCrLf & "w - Show Audio Waves" & vbCrLf & "left/right - seek backward/forward 10 seconds" & vbCrLf & "down/up - seek backward/forward 1 minute" & vbCrLf & "page down/page up - seek backward/forward 10 minutes" & vbCrLf & "mouse click - seek to percentage in file corresponding to fraction of width"
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    TerminateProcess OpenProcess(1&, True, ffplayPID), 0&
End Sub

Private Sub lvwPlayList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvwPlayList.ListItems.Count = 0 Then Exit Sub
    If Button = 2 And lvwPlayList.SelectedItem <> "" Then
        
        PopupMenu mnuPop
    End If
End Sub

Private Sub lvwPlayList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i&
    Dim lx As ListItem
    
    If Data.GetFormat(vbCFFiles) Then
        ReDim s(1 To Data.Files.Count)
        For i = 1 To Data.Files.Count
            If PathIsDirectory(Data.Files(i)) <> FILE_ATTRIBUTE_DIRECTORY Then
                DoEvents
                Set lx = lvwPlayList.ListItems.Add(, , Data.Files(i))
                
                If SupportOrNot(FsoSys.GetExtensionName(Data.Files(i))) Then
                    lx.SubItems(2) = FsoSys.GetExtensionName(Data.Files(i))
                Else
                    lx.SubItems(2) = "Not Supported"
                End If
            End If
        Next
    End If
End Sub


Private Sub mnuClearAll_Click()
    lvwPlayList.ListItems.Clear
End Sub

Private Sub mnuDelSelected_Click()
    With lvwPlayList
        Dim i As Long
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Selected = True Then
                .ListItems.Remove i
            End If
        Next
    End With
    
End Sub

Private Sub mnuFullScreen_Click()
    mnuFullScreen.Checked = Not mnuFullScreen.Checked
    WriteKey "Player", "FullScreen", mnuFullScreen.Checked, App.Path & "\YoutubeGrabberOption.ini"
    
End Sub

Private Sub mnuNoAudio_Click()
    mnuNoAudio.Checked = Not mnuNoAudio.Checked
    WriteKey "Player", "NoAudio", mnuNoAudio.Checked, App.Path & "\YoutubeGrabberOption.ini"
    
End Sub

Private Sub mnuNoVideo_Click()
    mnuNoVideo.Checked = Not mnuNoVideo.Checked
    WriteKey "Player", "NoVideo", mnuNoVideo.Checked, App.Path & "\YoutubeGrabberOption.ini"
    
End Sub

Private Sub mnuPlaySelected_Click()
    Play
End Sub

Private Sub mnuSFile_Click()
    frmSupportedFileType.Show 1
End Sub

Private Sub mnuSMRDFT_Click()
    mnuSMRDFT.Checked = Not mnuSMRDFT.Checked
    If mnuSMRDFT.Checked Then
        WriteKey "Player", "ShowMode", 2, App.Path & "\YoutubeGrabberOption.ini"
    End If
    mnuSMVideo.Checked = False
    mnuSMWaves.Checked = False
End Sub

Private Sub mnuSMVideo_Click()
    mnuSMVideo.Checked = Not mnuSMVideo.Checked
    If mnuSMVideo.Checked Then
        WriteKey "Player", "ShowMode", 0, App.Path & "\YoutubeGrabberOption.ini"
    End If
    mnuSMRDFT.Checked = False
    mnuSMWaves.Checked = False
End Sub

Private Sub mnuSMWaves_Click()
    mnuSMWaves.Checked = Not mnuSMWaves.Checked
    If mnuSMWaves.Checked Then
        WriteKey "Player", "ShowMode", 1, App.Path & "\YoutubeGrabberOption.ini"
    End If
    mnuSMRDFT.Checked = False
    mnuSMVideo.Checked = False
End Sub

Private Function SupportOrNot(ByVal TestFileName As String) As Boolean
    If InStr(codecs, "|/\<>*." & TestFileName) > 0 Then
        SupportOrNot = True
    Else
        SupportOrNot = False
    End If
End Function


Private Sub txtLink_Change()
    If Len(txtLink.Text) > 0 Then
        cmdAddNetworkLink.Enabled = True
    Else
        cmdAddNetworkLink.Enabled = False
    End If
    
End Sub

Private Sub Play()
    cmdExit.Enabled = True
    Dim PlayPara As String
    If mnuFullScreen.Checked = True Then PlayPara = PlayPara & " -fs"
    If mnuNoAudio.Checked = True Then PlayPara = PlayPara & " -an"
    If mnuNoVideo.Checked = True Then PlayPara = PlayPara & " -vn"
    If mnuSMVideo.Checked = True Then
        PlayPara = PlayPara & " -showmode 0"
    ElseIf mnuSMWaves.Checked = True Then
        PlayPara = PlayPara & " -showmode 1"
    ElseIf mnuSMRDFT.Checked = True Then
        PlayPara = PlayPara & " -showmode 2"
    End If
    lvwPlayList.SelectedItem.SubItems(1) = "Playing"
    WaitExeRun App.Path & "\ffplay.exe " & """" & lvwPlayList.SelectedItem & """" & PlayPara
    lvwPlayList.SelectedItem.SubItems(1) = ""
    cmdExit.Enabled = False
    
End Sub
