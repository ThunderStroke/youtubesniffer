VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG Youtube Video Downloader"
   ClientHeight    =   8580
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Downlaod List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   7335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   7095
         TabIndex        =   19
         Top             =   3000
         Width           =   7095
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear All"
            Height          =   495
            Left            =   0
            TabIndex        =   24
            Top             =   960
            Width           =   3615
         End
         Begin VB.CommandButton cmdDownSelected 
            Caption         =   "Download Selected"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   7095
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Selected"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3600
            TabIndex        =   22
            Top             =   960
            Width           =   3495
         End
         Begin VB.CommandButton cmdDownAll 
            Caption         =   "Start Download All"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   3615
         End
         Begin VB.CommandButton cmdStopAllDownload 
            Caption         =   "Stop Download All"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3600
            TabIndex        =   20
            Top             =   480
            Width           =   3495
         End
      End
      Begin MSComctlLib.ListView lvwDownload 
         Height          =   2655
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":5F32
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Download Link"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Quality"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Title"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Format"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   5040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSelectFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox DownTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   5655
      End
      Begin Project1.Xp_ProgressBar myProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
      End
      Begin VB.Label lblTime 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   2400
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Time Taken:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label DownState 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   1560
         Width           =   7215
      End
      Begin VB.Label LbMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Download Status:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label DownSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Throughput 0.00 KB/s"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   7215
      End
      Begin VB.Label LbMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Download Speed:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label LbMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Donwload Link:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label DownForm 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label LbMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Save To:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LbMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Download Progress:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Downloading 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0 bytes transferred (0%)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   7215
      End
   End
   Begin VB.CommandButton blnStop 
      Caption         =   "Stop Download"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton blnStart 
      Caption         =   "Start Download"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuProgressStyle 
         Caption         =   "Progress Bar Style"
         Begin VB.Menu mnuProDef 
            Caption         =   "Default"
         End
         Begin VB.Menu mnuProBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuProDBlue 
            Caption         =   "Dark Blue"
         End
         Begin VB.Menu mnuProGold 
            Caption         =   "Gold"
         End
         Begin VB.Menu mnuProGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuProGrey 
            Caption         =   "Grey"
         End
         Begin VB.Menu mnuProOrange 
            Caption         =   "Orange"
         End
         Begin VB.Menu mnuProRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuProCycle 
            Caption         =   "Cycle"
         End
         Begin VB.Menu mnuProRan 
            Caption         =   "Random"
         End
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Auto Convert Video"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open Folder After Download"
      End
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents wd As WininetDown
Attribute wd.VB_VarHelpID = -1
'Dim IsDownloading As Boolean
Dim DownloadedOrCanceled As Boolean
Dim NotClick As Boolean

Private Sub blnStart_Click()
    'IsDownloading = True
    If NotClick = False Then
        lvwDownload.SelectedItem.SubItems(2) = "Downloading..."
        If frmMain.lvwDownloadList.ListItems.Count <> 0 Then frmMain.lvwDownloadList.ListItems.Item(lvwDownload.SelectedItem.Index).SubItems(2) = "Downloading..."
        
    End If
    blnStart.Enabled = False
    blnStop.Enabled = True
    DownTo.Locked = True
    cmdDelete.Enabled = False
    cmdClear.Enabled = False
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    
    mnuOption.Enabled = False
    
    wd.URL = DownForm.Caption
    
    
    wd.filename = DownTo.Text
    
    'DownTo.Text = wd.FileName
    
    If wd.filename <> "" Then
        DownTo.Text = wd.filename
    Else
        'Dim DownX As Long, dx As Long
        'For dx = 1 To lvwDownload.ListItems.Count
        'If lvwDownload.ListItems.Item(dx).SubItems(2) = "Pending..." Then
        'DownX = dx
        'End If
        
        'Next
        'If DownX = 0 Then DownX = 1
        If DownForm.Caption = "" Then
            DownForm.Caption = lvwDownload.ListItems.Item(1)
        End If
        DownTo.Text = App.Path & "\" & lvwDownload.ListItems.Item(1).SubItems(3) & Right(lvwDownload.ListItems.Item(1).SubItems(4), Len(lvwDownload.ListItems.Item(1).SubItems(4)) - 1)
        wd.filename = DownTo.Text
    End If
    
    DownloadedOrCanceled = wd.DownloadToFile
    'IsDownloading = DownloadedOrCanceled
    If NotClick = False Then
        lvwDownload.SelectedItem.SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
        If frmMain.lvwDownloadList.ListItems.Count <> 0 Then frmMain.lvwDownloadList.ListItems.Item(lvwDownload.SelectedItem.Index).SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
        
    End If
    'If DownloadedOrCanceled = True And mnuConvert.Checked = True Then
    'frmConvert.txtIn.Text = DownTo.Text
    'frmConvert.Show 1
    'End If
    
    If mnuOpenFolder.Checked = True Then
        Shell "explorer /select," & DownTo.Text
    End If
    DownTo.Locked = False
    mnuOption.Enabled = True
    cmdDelete.Enabled = True
    cmdClear.Enabled = True
    cmdDownAll.Enabled = True
    cmdDownSelected.Enabled = True
End Sub

Private Sub blnStop_Click()
    myProgressBar1.Value = 0
    blnStop.Enabled = False
    blnStart.Enabled = True
    wd.bStop
End Sub



Private Sub cmdClear_Click()
    lvwDownload.ListItems.Clear
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    With lvwDownload
        Dim i As Long
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Selected = True Then
                .ListItems.Remove i
            End If
        Next
    End With
End Sub

Private Sub cmdDownAll_Click()
    Dim i As Long
    cmdStopAllDownload.Enabled = True
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    lvwDownload.Enabled = False
    For i = 1 To lvwDownload.ListItems.Count
        DownTo.Text = App.Path & "\" & lvwDownload.ListItems.Item(i).SubItems(3) & Right(lvwDownload.ListItems.Item(i).SubItems(4), Len(lvwDownload.ListItems.Item(i).SubItems(4)) - 1)
        DownForm.Caption = lvwDownload.ListItems.Item(i)
        NotClick = True
        lvwDownload.ListItems.Item(i).SubItems(2) = "Downloading..."
        frmMain.lvwDownloadList.ListItems.Item(lvwDownload.SelectedItem.Index).SubItems(2) = "Downloading..."
        
        blnStart_Click
        lvwDownload.ListItems.Item(i).SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
        frmMain.lvwDownloadList.ListItems.Item(lvwDownload.SelectedItem.Index).SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
        
        If DownloadedOrCanceled = False Then Exit For
        
        DoEvents
        
    Next
    lvwDownload.Enabled = True
    NotClick = False
    cmdStopAllDownload.Enabled = False
    cmdDownAll.Enabled = True
    cmdDownSelected.Enabled = True
End Sub

Private Sub cmdDownSelected_Click()
    lvwDownload.Enabled = False
    cmdDownAll.Enabled = False
    cmdClear.Enabled = False
    cmdDelete.Enabled = False
    DownForm.Caption = lvwDownload.SelectedItem
    DownTo.Text = App.Path & "\" & lvwDownload.SelectedItem.SubItems(3) & Right(lvwDownload.SelectedItem.SubItems(4), Len(lvwDownload.SelectedItem.SubItems(4)) - 1)
    
    blnStart_Click
    lvwDownload.SelectedItem.SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
    frmMain.lvwDownloadList.ListItems.Item(lvwDownload.SelectedItem.Index).SubItems(2) = IIf(DownloadedOrCanceled, "Downloaded", "Download Failed")
    
    lvwDownload.Enabled = True
    cmdDownAll.Enabled = True
    cmdClear.Enabled = True
    cmdDelete.Enabled = True
End Sub

Private Sub cmdSelectFile_Click()
    With cdFile
        .Filter = frmMain.txtCodec.Text & " (" & frmMain.txtExtension.Text & ")" & "|" & frmMain.txtExtension.Text
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        Else
            
            DownTo.Text = .filename
            If Dir$(DownTo.Text) <> "" Then
                Static Num As Integer
                Dim TextToTmp As String
                If TextToTmp = "" Then TextToTmp = DownTo.Text
                Do Until Dir$(DownTo.Text) = ""
                    If Num = 0 Then Num = 1
                    DownTo.Text = Mid$(TextToTmp, 1, InStrRev(TextToTmp, ".") - 1) & "(" & Num & ")" & Right$(TextToTmp, Len(TextToTmp) - InStrRev(TextToTmp, ".") + 1)
                    wd.filename = DownTo.Text
                    Num = Num + 1
                Loop
            End If
        End If
        
        
    End With
End Sub


Private Sub cmdStopAllDownload_Click()
    cmdDownAll.Enabled = True
    cmdStopAllDownload.Enabled = False
    cmdDownSelected.Enabled = True
    lvwDownload.Enabled = True
    blnStop_Click
End Sub



Private Sub Form_Load()
    
    Set wd = New WininetDown
    wd.URL = DownForm.Caption
    
    'wd.FileName = App.Path & "\" & frmMain.txtTitle.Text & Right(frmMain.txtExtension.Text, Len(frmMain.txtExtension.Text) - 1)
    wd.filename = DownTo.Text
    mnuOpenFolder.Checked = GetIni("Downloader", "AutoConvertVideo", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuConvert.Checked = GetIni("Downloader", "AutoOpenFolder", True, App.Path & "\YoutubeGrabberOption.ini")
    If Val(GetIni("Downloader", "ProgressBarStyle", XP_Default, App.Path & "\YoutubeGrabberOption.ini")) < 8 Then
        myProgressBar1.ProgressLook = Val(GetIni("Downloader", "ProgressBarStyle", XP_Default, App.Path & "\YoutubeGrabberOption.ini"))
        Select Case myProgressBar1.ProgressLook
        Case XP_Default
            mnuProDef.Checked = True
        Case XP_DarkBlue
            mnuProDBlue.Checked = True
        Case XP_Gold
            mnuProGold.Checked = True
        Case XP_Green
            mnuProGreen.Checked = True
        Case XP_Grey
            mnuProGrey.Checked = True
        Case XP_Orange
            mnuProOrange.Checked = True
        Case XP_Red
            mnuProRed.Checked = True
        Case XP_Blue
            mnuProBlue.Checked = True
        End Select
    ElseIf Val(GetIni("Downloader", "ProgressBarStyle", XP_Default, App.Path & "\YoutubeGrabberOption.ini")) >= 8 Then
        Select Case Val(GetIni("Downloader", "ProgressBarStyle", XP_Default, App.Path & "\YoutubeGrabberOption.ini"))
        Case 8
            mnuProCycle.Checked = True
        Case 9
            mnuProRan.Checked = True
        End Select
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
End Sub



Private Sub lvwDownload_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelete.Enabled = True
    cmdDownSelected.Enabled = True
    cmdDownAll.Enabled = True
End Sub

Private Sub mnuConvert_Click()
    If mnuConvert.Checked = True Then
        mnuConvert.Checked = False
        WriteKey "Downloader", "AutoConvertVideo", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuConvert.Checked = True
        WriteKey "Downloader", "AutoConvertVideo", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

Private Sub mnuOpenFolder_Click()
    If mnuOpenFolder.Checked = True Then
        mnuOpenFolder.Checked = False
        WriteKey "Downloader", "AutoOpenFolder", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuOpenFolder.Checked = True
        WriteKey "Downloader", "AutoOpenFolder", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

Private Sub mnuProBlue_Click()
    If mnuProBlue.Checked = True Then
        mnuProBlue.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        myProgressBar1.ProgressLook = XP_Blue
        WriteKey "Downloader", "ProgressBarStyle", XP_Blue, App.Path & "\YoutubeGrabberOption.ini"
        mnuProBlue.Checked = True
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
    End If
End Sub

Private Sub mnuProCycle_Click()
    If mnuProCycle.Checked = True Then
        mnuProCycle.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        mnuProCycle.Checked = True
        WriteKey "Downloader", "ProgressBarStyle", 8, App.Path & "\YoutubeGrabberOption.ini"
        
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
    End If
End Sub

Private Sub mnuProDBlue_Click()
    If mnuProDBlue.Checked = True Then
        mnuProDBlue.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        mnuProDBlue.Checked = True
        WriteKey "Downloader", "ProgressBarStyle", XP_DarkBlue, App.Path & "\YoutubeGrabberOption.ini"
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
        myProgressBar1.ProgressLook = XP_DarkBlue
    End If
End Sub

Private Sub mnuProDef_Click()
    
    If mnuProDef.Checked = True Then
        mnuProDef.Checked = False
    Else
        mnuProDef.Checked = True
        myProgressBar1.ProgressLook = XP_Default
        WriteKey "Downloader", "ProgressBarStyle", XP_Default, App.Path & "\YoutubeGrabberOption.ini"
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
    End If
End Sub

Private Sub mnuProGold_Click()
    If mnuProGold.Checked = True Then
        mnuProGold.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        mnuProGold.Checked = True
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
        myProgressBar1.ProgressLook = XP_Gold
        WriteKey "Downloader", "ProgressBarStyle", XP_Gold, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

Private Sub mnuProGreen_Click()
    If mnuProGreen.Checked = True Then
        mnuProGreen.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        mnuProGreen.Checked = True
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
        myProgressBar1.ProgressLook = XP_Green
        WriteKey "Downloader", "ProgressBarStyle", XP_Green, App.Path & "\YoutubeGrabberOption.ini"
        
    End If
End Sub

Private Sub mnuProGrey_Click()
    If mnuProGrey.Checked = True Then
        mnuProGrey.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        mnuProGrey.Checked = True
        myProgressBar1.ProgressLook = XP_Grey
        WriteKey "Downloader", "ProgressBarStyle", XP_Grey, App.Path & "\YoutubeGrabberOption.ini"
        
    End If
End Sub

Private Sub mnuProOrange_Click()
    If mnuProOrange.Checked = True Then
        mnuProOrange.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        mnuProOrange.Checked = True
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
        myProgressBar1.ProgressLook = XP_Orange
        WriteKey "Downloader", "ProgressBarStyle", XP_Orange, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

Private Sub mnuProRan_Click()
    If mnuProRan.Checked = True Then
        mnuProRan.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        mnuProRan.Checked = True
        WriteKey "Downloader", "ProgressBarStyle", 9, App.Path & "\YoutubeGrabberOption.ini"
        
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        If mnuProRed.Checked = True Then mnuProRed.Checked = False
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
    End If
End Sub

Private Sub mnuProRed_Click()
    If mnuProRed.Checked = True Then
        mnuProRed.Checked = False
        
        
    Else
        If mnuProDef.Checked = True Then mnuProDef.Checked = False
        If mnuProBlue.Checked = True Then mnuProBlue.Checked = False
        If mnuProDBlue.Checked = True Then mnuProDBlue.Checked = False
        If mnuProGold.Checked = True Then mnuProGold.Checked = False
        If mnuProGreen.Checked = True Then mnuProGreen.Checked = False
        If mnuProRan.Checked = True Then mnuProRan.Checked = False
        If mnuProCycle.Checked = True Then mnuProCycle.Checked = False
        mnuProRed.Checked = True
        If mnuProOrange.Checked = True Then mnuProOrange.Checked = False
        If mnuProGrey.Checked = True Then mnuProGrey.Checked = False
        myProgressBar1.ProgressLook = XP_Red
        WriteKey "Downloader", "ProgressBarStyle", XP_Red, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub





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
        If mnuProCycle.Checked = True Then
            myProgressBar1.ProgressLook = BShare Mod 8
        ElseIf mnuProRan.Checked = True Then
            myProgressBar1.ProgressLook = (9 + Int(Rnd() * 10)) Mod 8
        End If
        
        
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
    
    DownState.Caption = strMsg
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
    
    DownState.Caption = strMsg
    blnStart.Enabled = True
    blnStop.Enabled = False
End Sub

