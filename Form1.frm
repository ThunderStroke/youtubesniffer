VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG Youtube Video Grabber"
   ClientHeight    =   7935
   ClientLeft      =   1680
   ClientTop       =   3495
   ClientWidth     =   18825
   BeginProperty Font 
      Name            =   "Segoe UI"
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
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1255
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   42
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   6975
      Left            =   15960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox lstState 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "Program Status"
      Top             =   6720
      Width           =   15615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Download List"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   10920
      TabIndex        =   29
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdDownSelected 
         Caption         =   "Download Select"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Download Selected Video"
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CommandButton cmdDownAll 
         Caption         =   "Download All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   39
         ToolTipText     =   "Download All Video In The Download List"
         Top             =   5520
         Width           =   2175
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Selected"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         ToolTipText     =   "Delect Selected Video Entry"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Clear All Entries In The Download List"
         Top             =   5880
         Width           =   2415
      End
      Begin MSComctlLib.ListView lvwDownloadList 
         Height          =   5175
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Download List"
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   9128
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Format"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Click On The Node Select Quality"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   28
      Top             =   2160
      Width           =   10215
      Begin MSComctlLib.TreeView tvwQuality 
         Height          =   1815
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Click At The Link To Select A Download Link"
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3201
         _Version        =   393217
         LabelEdit       =   1
         Style           =   4
         FullRowSelect   =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":6094
      End
   End
   Begin AGYouTubeVideoGrabber.Tray Tray 
      Left            =   120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      PictureIcon     =   "Form1.frx":61F6
   End
   Begin VB.Timer trmGetClipData 
      Interval        =   100
      Left            =   0
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetFileSize 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdVisit 
         Caption         =   "Visit"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8160
         TabIndex        =   36
         ToolTipText     =   "Visit YouTube Channel"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtLength 
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Video's Duration"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtView 
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Video's Views"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtLink 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Enter YouTube Video Link Here"
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "YouTube Video ID"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Video's Title"
         Top             =   960
         Width           =   7335
      End
      Begin VB.TextBox txtHome 
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Author's YouTube Channel"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtUploader 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Video's Author"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Video Length:"
         Height          =   255
         Left            =   5760
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Views:"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Youtube Link:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Youtube ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Video's Title: "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Uploader's Channel:"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Hit Enter When You Are Done"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   7680
         TabIndex        =   9
         ToolTipText     =   "Hit Enter When You're Done"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Uploaded By:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download Option"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   10695
      Begin VB.CommandButton cmdAddDown 
         Caption         =   "Add To Download List >>>"
         Enabled         =   0   'False
         Height          =   975
         Left            =   7320
         TabIndex        =   35
         ToolTipText     =   "Add Selected Video To Download List "
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Open Selected Video In Browser"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Open The Video In Your Default Browser"
         Top             =   3720
         Width           =   7095
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Download Selected Video"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "Download The Selected Video"
         Top             =   3240
         Width           =   7095
      End
      Begin VB.TextBox txtQuality 
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Video's Quality"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtFileSize 
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Video File Size"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtExtension 
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Video's Format"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtCodec 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Video's Codec"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtDownloadLink 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Selected Download Liink"
         Top             =   2520
         Width           =   8655
      End
      Begin VB.Label Label15 
         Caption         =   "Quality:"
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "File Size: "
         Height          =   255
         Left            =   7560
         TabIndex        =   22
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "File Format:"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Codec:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Download Link:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSavePic 
         Caption         =   "Save Picture"
         Begin VB.Menu mnuSubSavePicAs 
            Caption         =   "Save Picture As..."
         End
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert An Existing Video"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyLink 
         Caption         =   "Copy Selected Download Link"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Option"
      Begin VB.Menu mnuGeneral 
         Caption         =   "Genaral"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMinToTray 
         Caption         =   "Minimize To Tray"
      End
      Begin VB.Menu munDashA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Main"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetLinkClip 
         Caption         =   "Automatically get link from clipboard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAppActivate 
         Caption         =   "Activate app when got link from clipboard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowStat 
         Caption         =   "Show Program Status Window"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart The Program"
      End
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "Check For Updates"
      Begin VB.Menu mnuCheckUpAuto 
         Caption         =   "Check Update When Program Starts"
      End
      Begin VB.Menu mnuCheckUpdateNow 
         Caption         =   "Check Updates Now"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuPopDL 
      Caption         =   "PopUpD"
      Visible         =   0   'False
      Begin VB.Menu mnuDownSelected 
         Caption         =   "Download Selected"
      End
      Begin VB.Menu mnuDownloadAll 
         Caption         =   "Download All"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
      Begin VB.Menu mnuDelSelected 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuLvwCopy 
         Caption         =   "Copy Selected Download Link"
      End
   End
   Begin VB.Menu mnuPopQua 
      Caption         =   "PopUpQua"
      Visible         =   0   'False
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "Collapse All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Pm As POINTAPI                                                              'Mouse Cursor Coordinate
'Dim PRct As RECT

'Dim OldFrmSize As POINTAPI                                                      ' Old Form Size
'Dim YoutubeDLink As String, WebSwf As String
Dim WebHtml As String
'DecodeString As String,
'Public DlFileEx As String
Dim URLFileSize As String                                                       'For Getting URL size True=Get
'Dim URLCheckTmp As Boolean
'Dim InetFail As Boolean
'Dim PrevWndProc_frmmain As Long


'Private Sub cmbDownloadOption_Click()
'If cmbDownloadOption.ListIndex = 0 Then
'txtDownloadLink.Text = ""
'txtCodec.Text = ""
'txtExtension.Text = ""
'txtFileSize.Text = ""
'cmdBrowse.Enabled = False
'cmdDown.Enabled = False
'txtQuality.Text = ""
'cmdAddDown.Enabled = False
'cmdPlay.Enabled = False
'Exit Sub
'End If
'MsgBox cmbDownloadOption.List(0)
'txtDownloadLink.Text = cmbDownloadOption.List(cmbDownloadOption.ListIndex)
'End Sub

Private Sub cmdAddDown_Click()
    Dim x As ListItem
    
    Set x = lvwDownloadList.ListItems.Add(, , txtDownloadLink.Text)
    
    x.SubItems(1) = txtQuality.Text
    x.SubItems(2) = "Pending..."
    x.SubItems(3) = txtTitle.Text
    x.SubItems(4) = txtExtension.Text
    If lvwDownloadList.ListItems.Count <> 0 Then cmdDownAll.Enabled = True
End Sub

Private Sub cmdBrowse_Click()
    VisitURL txtDownloadLink.Text
End Sub

Private Sub cmdClear_Click()
    lvwDownloadList.ListItems.Clear
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    With lvwDownloadList
        Dim I As Long
        For I = .ListItems.Count To 1 Step -1
            If .ListItems(I).Selected = True Then
                .ListItems.Remove I
            End If
        Next
    End With
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
'??:12.4275000000052
End Sub

Private Sub cmdDown_Click()
    Load frmDownload
    Dim x As ListItem
    Set x = frmDownload.lvwDownload.ListItems.Add(, , txtDownloadLink.Text)
    frmDownload.DownForm.Caption = txtDownloadLink.Text
    frmDownload.DownTo.Text = App.Path & "\" & txtTitle.Text & Right(txtExtension.Text, Len(txtExtension.Text) - 1)
    'lvwDownloadList.SelectedItem
    x.SubItems(1) = txtQuality.Text
    x.SubItems(2) = "Pending..."
    x.SubItems(3) = txtTitle.Text
    x.SubItems(4) = txtExtension.Text
    
    
    
    frmDownload.Show 1
    'Unload frmDownload
End Sub

Private Sub cmdDownAll_Click()
    Dim I As Long
    Load frmDownload
    Dim x As ListItem
    For I = 1 To lvwDownloadList.ListItems.Count
        'Debug.Print lvwDownloadList.ListItems.Item(i)
        'Debug.Print lvwDownloadList.ListItems.Item(i).SubItems(1)
        Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.ListItems.Item(I))
        x.SubItems(1) = lvwDownloadList.ListItems.Item(I).SubItems(1)
        x.SubItems(2) = lvwDownloadList.ListItems.Item(I).SubItems(2)
        x.SubItems(3) = lvwDownloadList.ListItems.Item(I).SubItems(3)
        x.SubItems(4) = lvwDownloadList.ListItems.Item(I).SubItems(4)
    Next
    frmDownload.Show 1
End Sub

Private Sub cmdDownSelected_Click()
    Load frmDownload
    Dim x As ListItem
    Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.SelectedItem)
    frmDownload.DownForm.Caption = lvwDownloadList.SelectedItem
    frmDownload.DownTo.Text = App.Path & "\" & lvwDownloadList.SelectedItem.SubItems(3) & Right(lvwDownloadList.SelectedItem.SubItems(4), Len(lvwDownloadList.SelectedItem.SubItems(4)) - 1)
    'lvwDownloadList.SelectedItem
    x.SubItems(1) = lvwDownloadList.SelectedItem.SubItems(1)
    x.SubItems(2) = lvwDownloadList.SelectedItem.SubItems(2)
    x.SubItems(3) = lvwDownloadList.SelectedItem.SubItems(3)
    x.SubItems(4) = lvwDownloadList.SelectedItem.SubItems(4)
    frmDownload.Show 1
End Sub



'Private Sub cmdPlay_Click()
    'Load frmPlay
    'Dim lx As ListItem
    'Set lx = frmPlay.lvwPlayList.ListItems.Add(, , txtDownloadLink.Text)
    'lx.SubItems(2) = "Network File"
    'frmPlay.Show 1
'End Sub

Private Sub cmdVisit_Click()
    VisitURL (txtHome.Text)
End Sub

Private Sub Form_Activate()
    Static Opened As Integer
    If Opened = 0 Then
        Opened = 1
    Else
        Exit Sub
    End If
    Load frmUpdate
    frmUpdate.Show
    frmUpdate.SetFocus
End Sub

Private Sub Form_Load()
    'cmbDownloadOption.AddItem "Please select..."
    If App.LogMode = 0 Then
        'MsgBox "Don't Run At IDE Please!", vbInformation
        RemoveMenu GetSystemMenu(hWnd, 0), SC_CLOSE, MF_REMOVE
        'End
    ElseIf App.LogMode Then
        Attach Me.hWnd
    End If
    'If Dir$(App.Path & "\ffmpeg.exe") = "" Then
    'MsgBox "ffmpeg.exe not found" & vbNewLine & "Please Redownload!", vbCritical, "Error!"
    'End
    'End If
    'If Dir$(App.Path & "\BAK_" & App.EXEName & ".exe") <> "" Then
    'Kill App.Path & "\BAK_" & App.EXEName & ".exe"
    'End If
    'If Dir$(App.Path & "\ffplay.exe") = "" Then
    'MsgBox "ffplay.exe not found" & vbNewLine & "Please Redownload!", vbCritical, "Error!"
    'End
    'End If
    'With OldPicSize
        '.x = Picture2.Width
        '.y = Picture2.Height
    'End With
    'With OldFrmSize
    '.x = Me.Width
    '.y = Me.Height
    'End With
    
    'cmbDownloadOption.ListIndex = 0
    
    lstAdd "Loading User Settings"
    
    mnuMinToTray.Checked = GetIni("General", "MinimizeToTray", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuAppActivate.Checked = GetIni("Main", "ActivateApp", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuGetLinkClip.Checked = GetIni("Main", "AutoGetLink", True, App.Path & "\YoutubeGrabberOption.ini")
    mnuCheckUpAuto.Checked = GetIni("Main", "AutoUpdate", True, App.Path & "\YoutubeGrabberOption.ini")
    'mnuShowStat.Checked = GetIni("General", "ShowStatusWindow", True, App.Path & "\YoutubeGrabberOption.ini")
    lstAdd "User Settings Loaded"
    'If mnuShowStat.Checked Then Form2.Show
    'Form2.Width = Me.Width
    'Form2.lstState.Width = Me.Width
    'pWndProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    'SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WindowProc
    'PrevWndProc_frmmain = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    'Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf SubWndProc_frmmain)
    
    
    'lvwFullRow lvwDownloadList
    
    
    'Picture3.FontSize = 8
    'Picture3.ForeColor = vbRed
    'Picture3.Print "I Still Can't"
    'Picture3.Print "Fix This..."
    'Picture3.Print "ScreenShot"
    'Picture3.Print "Currently"
    'Picture3.Print "Unavailable."
    'Picture3.ForeColor = vbBlue
    'Picture3.Print "Sorry :("
    'Picture3.Print "   ---GaryNg"
    
    'lblVer.Caption = "Build: " & App.Major & "." & App.Minor & "." & App.Revision & "  - By GaryNg from AG Dev Team"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 And mnuMinToTray.Checked = True Then
        Me.Hide
        'Form2.Hide
        Tray.Show
        Tray.ShowBubble App.Title, "Click Me To Restore", NIIF_GUID
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SetWindowLong Me.hwnd, GWL_WNDPROC, pWndProc
    Tray.Hide
    
    If App.LogMode Then
        Detach Me.hWnd
    End If
    'Unload Form2
    'Unload frmConvert
    Unload frmDownload
    Unload frmMain
    'Unload frmPlay
    Unload frmAbout
    ' Unload frmSupportedFileType
    Unload frmUpdate
    Unload Me
    TerminateProcess GetCurrentProcess, ByVal 0&
    
End Sub

Private Sub InetFileSize_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Dim vtData As String
    If State = 12 Then
        vtData = InetFileSize.GetHeader
        Do While InetFileSize.StillExecuting
            DoEvents
        Loop
        Dim DataTmp As String
        DataTmp = Replace$(vtData, vbCrLf, "")
        Debug.Print DataTmp
        URLFileSize = VBStrFormatByteSize(Val(Mid$(DataTmp, InStr(LCase$(DataTmp), "content-length: ") + Len("content-length: "), InStr(LCase$(DataTmp), "connection: ") - InStr(LCase$(DataTmp), "content-length: ") - Len("content-length: "))))
    End If
End Sub
  
    Private Sub Label3_Click()
    txtLink_KeyDown 13, 1
End Sub

Private Sub lvwDownloadList_ItemClick(ByVal Item As ListItem)
    cmdDelete.Enabled = True
    cmdDownSelected.Enabled = True
    cmdDownAll.Enabled = True
End Sub

Private Sub lvwDownloadList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvwDownloadList.ListItems.Count = 0 Then Exit Sub
    If Button = 2 And lvwDownloadList.SelectedItem <> "" Then
        PopupMenu mnuPopDL
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAppActivate_Click()
    If mnuAppActivate.Checked = True Then
        mnuAppActivate.Checked = False
        WriteKey "Main", "ActivateApp", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuAppActivate.Checked = True
        WriteKey "Main", "ActivateApp", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
    
End Sub


Private Sub mnuCheckUpAuto_Click()
    mnuCheckUpAuto.Checked = Not mnuCheckUpAuto.Checked
    WriteKey "Main", "AutoUpdate", mnuCheckUpAuto.Checked, App.Path & "\YoutubeGrabberOption.ini"
End Sub

Private Sub mnuCheckUpdateNow_Click()
    Load frmUpdate
    frmUpdate.Show
End Sub

Private Sub mnuClearAll_Click()
    lvwDownloadList.ListItems.Clear
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub mnuCollapse_Click()
    Dim I As Long
    For I = 1 To tvwQuality.Nodes.Count
        tvwQuality.Nodes(I).Expanded = False
    Next I
End Sub

Private Sub mnuConvert_Click()
    'frmConvert.Show 1
End Sub

Private Sub mnuCopyLink_Click()
    If txtDownloadLink.Text <> "" Then
        Clipboard.Clear
        Clipboard.SetText txtDownloadLink.Text
        lstAdd "Download Link Copied."
        
    End If
End Sub

Private Sub mnuDelSelected_Click()
    With lvwDownloadList
        Dim I As Long
        For I = .ListItems.Count To 1 Step -1
            If .ListItems(I).Selected = True Then
                .ListItems.Remove I
            End If
        Next
    End With
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub mnuDownloadAll_Click()
    Dim I As Long
    Load frmDownload
    Dim x As ListItem
    For I = 1 To lvwDownloadList.ListItems.Count
        'Debug.Print lvwDownloadList.ListItems.Item(i)
        'Debug.Print lvwDownloadList.ListItems.Item(i).SubItems(1)
        Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.ListItems.Item(I))
        x.SubItems(1) = lvwDownloadList.ListItems.Item(I).SubItems(1)
        x.SubItems(2) = lvwDownloadList.ListItems.Item(I).SubItems(2)
        x.SubItems(3) = lvwDownloadList.ListItems.Item(I).SubItems(3)
        x.SubItems(4) = lvwDownloadList.ListItems.Item(I).SubItems(4)
    Next
    frmDownload.Show 1
End Sub

Private Sub mnuDownSelected_Click()
    Load frmDownload
    Dim x As ListItem
    Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.SelectedItem)
    frmDownload.DownForm.Caption = lvwDownloadList.SelectedItem
    frmDownload.DownTo.Text = App.Path & "\" & lvwDownloadList.SelectedItem.SubItems(3) & Right(lvwDownloadList.SelectedItem.SubItems(4), Len(lvwDownloadList.SelectedItem.SubItems(4)) - 1)
    'lvwDownloadList.SelectedItem
    x.SubItems(1) = lvwDownloadList.SelectedItem.SubItems(1)
    x.SubItems(2) = lvwDownloadList.SelectedItem.SubItems(2)
    x.SubItems(3) = lvwDownloadList.SelectedItem.SubItems(3)
    x.SubItems(4) = lvwDownloadList.SelectedItem.SubItems(4)
    frmDownload.Show 1
End Sub

Private Sub mnuExpand_Click()
    Dim I As Long
    For I = 1 To tvwQuality.Nodes.Count
        tvwQuality.Nodes(I).Expanded = True
    Next I
End Sub

Private Sub mnuGetLinkClip_Click()
    If mnuGetLinkClip.Checked = True Then
        mnuGetLinkClip.Checked = False
        WriteKey "Main", "AutoGetLink", False, App.Path & "\YoutubeGrabberOption.ini"
        trmGetClipData.Enabled = False
        'mnuAppActivate.Checked = False
    Else
        mnuGetLinkClip.Checked = True
        WriteKey "Main", "AutoGetLink", True, App.Path & "\YoutubeGrabberOption.ini"
        trmGetClipData.Enabled = True
        'mnuAppActivate.Checked = True
    End If
    
End Sub

Private Sub mnuLvwCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lvwDownloadList.SelectedItem.Text
    lstAdd "Link Copied"
End Sub

Private Sub mnuMinToTray_Click()
    If mnuMinToTray.Checked = True Then
        mnuMinToTray.Checked = False
        WriteKey "General", "MinimizeToTray", False, App.Path & "\YoutubeGrabberOption.ini"
    Else
        mnuMinToTray.Checked = True
        WriteKey "General", "MinimizeToTray", True, App.Path & "\YoutubeGrabberOption.ini"
    End If
End Sub

'Private Sub mnuPlayer_Click()
'frmPlay.Show 1
'End Sub

Private Sub mnuRestart_Click()
    Do Until OpenProcess(&H400, 0, Shell(App.Path & "\" & App.EXEName & ".exe")) <> 0
        DoEvents
    Loop
    Unload Me
End Sub

'Private Sub mnuShowStat_Click()
'mnuShowStat.Checked = Not mnuShowStat.Checked
'Form2.Visible = Not Form2.Visible
'WriteKey "General", "ShowStatusWindow", mnuShowStat.Checked, App.Path & "\YoutubeGrabberOption.ini"
'End Sub

Private Sub mnuSubSavePicAs_Click()
    With cdSave
        .DefaultExt = ".bmp"
        .Filter = "Bitmap File (*.BMP)|*.BMP"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        Else
            If Dir$(.FileName) <> "" Then
                Static Num As Integer
                Dim TextPicTmp As String
                If TextPicTmp = "" Then TextPicTmp = .FileName
                Do Until Dir$(.FileName) = ""
                    If Num = 0 Then Num = 1
                    .FileName = Mid$(TextPicTmp, 1, InStrRev(TextPicTmp, ".") - 1) & "(" & Num & ")" & Right$(TextPicTmp, Len(TextPicTmp) - InStrRev(TextPicTmp, ".") + 1)
                    Num = Num + 1
                Loop
            End If
            Picture1.Picture = Picture1.Image
            SavePicture Picture1.Picture, .FileName
            lstAdd "ScreenShot Saved."
            
        End If
    End With
End Sub

'Private Sub Picture2_Click()
'Static Run As Boolean
'If Run = False Then
'Picture2.Height = Picture1.Height + 200
' Me.Height = Picture1.Height + 1000
'Picture2.Width = Picture1.Width + 200
'Me.Width = Picture1.Width + 1000
'FitPictureToBox Picture1, Picture2
'Run = True
'ElseIf Run = True Then
'Picture2.Height = OldPicSize.y
'Picture2.Width = OldPicSize.x
'Me.Height = OldFrmSize.y
'Me.Width = OldFrmSize.X
'FitPictureToBox Picture1, Picture2
'Run = False
'End If
'End Sub

'Private Sub optDownload_Click(Index As Integer)
'If txtDownload(Index).Text <> "" Then
'Load frmDownload
'Select Case Index
'Case Is <= 1
'DlFileEx = ".flv"
'Case 2
'DlFileEx = ".mp4"
'End Select
'If frmDownload.Visible = False Then
'frmDownload.DownForm = txtDownload(Index).Text
'Else
'Dim NewDownForm As New frmDownload
'Load NewDownForm
'With NewDownForm
'.DownForm = txtDownload(Index).Text
'.Show
'End With
'End If
'cmdDown.Enabled = True
'cmdBrowse.Enabled = True
'URLForVisit = txtDownload(Index).Text

'Else
'cmdDown.Enabled = False
'cmdBrowse.Enabled = True
'End If

'End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuSavePic
    End If
End Sub

'Public Sub LoadWebImage(ByVal PicSource As PictureBox, PicDestination As PictureBox)
'On Error Resume Next
'Dim byt() As Byte
'Inet1.Cancel
'byt() = Inet1.OpenURL(Mid$(WebHtml, InStr(WebHtml, "<meta property=""og:image"" content=") + 35, InStr(WebHtml, "hqdefault.jpg"">") - InStr(WebHtml, "<meta property=""og:image"" content=") - 35 + 13), icByteArray) '这是图片的地址


'Set PicSource.Picture = ShowPictureFromStream("http://www.vbgood.com/vucenter/data/avatar/000/14/08/45_avatar_middle.jpg")
'FitPictureToBox PicSource, PicDestination

'End Sub

Private Sub Tray_Click()
    Me.WindowState = 0
    Me.Show
    'Form2.Show
    Tray.Hide
End Sub

Private Sub trmGetClipData_Timer()
    On Error Resume Next
    If mnuGetLinkClip.Checked = False Then Exit Sub
    If ExtractMatch(Clipboard.GetText, "v=([A-Za-z0-9-_]+)") = "" Then Exit Sub
    If InStr(LCase$(Clipboard.GetText), "www.youtube.com") <> 0 And LCase$(txtLink.Text) <> LCase$(Clipboard.GetText) Then
        lstAdd "Found Link In Clipboard"
        txtLink.Text = Clipboard.GetText
        If mnuAppActivate.Checked = True Then
            AppActivate Me.Caption
        End If
        txtLink_KeyDown 13, 0
    End If
    Exit Sub
ClipErr:
    lstAdd Err.Description & Err.Source
End Sub

Private Sub tvwQuality_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And tvwQuality.Nodes.Count > 0 Then
        PopupMenu mnuPopQua
    End If
End Sub

Private Sub tvwQuality_NodeClick(ByVal Node As Node)
    If tvwQuality.SelectedItem = txtDownloadLink.Text Then Exit Sub
    Dim I As Long
    If tvwQuality.SelectedItem.children <> 0 Then Exit Sub
    'If tvwQuality.SelectedItem.Child.Selected = False Then Exit Sub
    'If tvwQuality.SelectedItem.Child.Parent.Selected Then Exit Sub
    
    If tvwQuality.SelectedItem.children = 0 Then
        
        For I = 1 To tvwQuality.Nodes.Count
            If tvwQuality.Nodes(I).Selected Then
                lstAdd "Quality Selected. Analyzing"
                
                txtDownloadLink.Text = tvwQuality.Nodes(I).Text
                'Dim h As Long
                'For h = 0 To cmbDownloadOption.ListCount - 1
                'If cmbDownloadOption.List(h) = tvwQuality.Nodes(i).Text Then
                'cmbDownloadOption.ListIndex = h
                'Exit For
                'End If
                '    Next
                'cmbDownloadOption.List(cmbDownloadOption.ListIndex) = tvwQuality.Nodes(i).Text
            End If
        Next I
    End If
End Sub

Private Sub txtDownloadLink_Change()
    On Error GoTo goterr
    If txtDownloadLink.Text = "" Then Exit Sub
    lstAdd "Getting Video Info"
    tvwQuality.Enabled = False
    If InStr(txtDownloadLink.Text, "type=video/webm") Then
        txtCodec.Text = "video/webm"
        txtExtension.Text = "*.webm"
    ElseIf InStr(txtDownloadLink.Text, "&type=video/mp4") Then
        txtCodec.Text = "video/mp4"
        txtExtension.Text = "*.mp4"
    ElseIf InStr(txtDownloadLink.Text, "&type=video/x-flv") Then
        txtCodec.Text = "video/x-flv"
        txtExtension.Text = "*.flv"
    End If
    lstAdd "Video Info Loaded"
    
    lstAdd "Getting Video Quality"
    
    txtQuality.Text = ExtractMatch(txtDownloadLink.Text, "&quality=([a-zA-Z0-9]*)")
    lstAdd "Video Quality Loaded"
    
    
    'cmbDownloadOption.Enabled = False
    lstAdd "Getting Video File Size"
    
    InetFileSize.Cancel
    InetFileSize.Execute txtDownloadLink.Text, "GET"
    Do Until URLFileSize <> ""
        DoEvents
    Loop
    'cmbDownloadOption.Enabled = True
    
    txtFileSize.Text = URLFileSize
    URLFileSize = ""
    lstAdd "Video File Size Loaded"
    
    cmdBrowse.Enabled = True
    cmdDown.Enabled = True
    cmdAddDown.Enabled = True
    tvwQuality.Enabled = True
    'cmdPlay.Enabled = True
    Load frmDownload
    frmDownload.DownForm.Caption = txtDownloadLink.Text
    'cmbDownloadOption.SetFocus
    lstAdd "Everything Done"
    
    Exit Sub
goterr:
    lstAdd Err.Description
    
    'cmbDownloadOption.Enabled = True
    txtDownloadLink.Text = ""
    txtCodec.Text = ""
    txtExtension.Text = ""
    InetFileSize.Cancel
    Unload frmDownload
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    If KeyCode = 13 Then
        If txtLink.Text = "" Then CleanUp: Exit Sub
        txtLink.Enabled = False
        tvwQuality.SetFocus
        Dim x As Integer
        CleanUp
        If InStr(LCase(txtLink.Text), "www.youtube.com") = 0 Then lstAdd "Invalid Link": txtLink.Text = "": txtLink.SetFocus: Exit Sub
        
        lstAdd "Getting Web Source Code"
        If SeperateSWF(txtLink.Text) = "Hey! No data recieved! Check your network connection" Then
            lstAdd "Hey! No data recieved! Check your network connection"
            CleanUp
            Exit Sub
        End If
        lstAdd "Loading Video's Info"
        LoadVideoInfo txtLink.Text, txtID, txtTitle, txtView, txtLength, txtUploader, txtHome, Picture1, Picture2
        Dim strDownloadLinks() As String
        lstAdd "Processing Download Links"
        ProcessDownloadLinks strDownloadLinks
        Dim tvwNode As Node
        lstAdd "Analyzing And Catergorizing Video Quality"
        
        Dim strQuality As String
        Dim nI As Long
        For nI = 1 To UBound(strDownloadLinks)
            strQuality = LCase(ExtractMatch(strDownloadLinks(nI), "&quality=([a-zA-Z0-9]*)"))
            If strQuality <> "" Then
                Set tvwNode = tvwQuality.Nodes.Add(, , strQuality, strQuality)
                Set tvwNode = tvwQuality.Nodes.Add(strQuality, tvwChild, , strDownloadLinks(nI))
            End If
            DoEvents
        Next
        lstAdd "Everything Done."
        Exit Sub
    Else
        txtLink.SetFocus
    End If
    Exit Sub
goterr:
    lstAdd "Program Error!"
    CleanUp
    txtLink.Enabled = True
End Sub

Private Sub CleanUp()
    lstAdd "Reseting.."
    'cmbDownloadOption.Clear
    'cmbDownloadOption.AddItem "Please Select..."
    'cmbDownloadOption.ListIndex = 0
    txtID.Text = ""
    'txtLike.Text = ""
    'txtDislikes.Text = ""
    txtView.Text = ""
    tvwQuality.Nodes.Clear
    txtLink.Enabled = True
    txtTitle.Text = ""
    txtLength.Text = ""
    txtUploader.Text = ""
    txtHome.Text = ""
    txtDownloadLink.Text = ""
    txtCodec.Text = ""
    txtExtension.Text = ""
    txtQuality.Text = ""
    txtFileSize.Text = ""
    'txtLink.Text = ""
    'txtLink.Enabled = True
    Picture1 = LoadPicture(Null)
    Picture2 = LoadPicture(Null)
End Sub

Private Sub lstAdd(ByVal InName As String)
    lstState.AddItem Date & "-" & Time & ": " & InName
    lstState.ListIndex = lstState.NewIndex
End Sub
