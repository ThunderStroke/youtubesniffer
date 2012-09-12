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
   ClientWidth     =   15915
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
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1061
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstState 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "Program Status"
      Top             =   6720
      Width           =   15615
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   32
      Top             =   120
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet InetUpdate 
      Left            =   0
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Download List"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   40
         Top             =   5520
         Width           =   4575
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   0
            TabIndex        =   44
            ToolTipText     =   "Clear All Entries In The Download List"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Selected"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   43
            ToolTipText     =   "Delect Selected Video Entry"
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdDownAll 
            Caption         =   "Download All"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   42
            ToolTipText     =   "Download All Video In The Download List"
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton cmdDownSelected 
            Caption         =   "Download Select"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   41
            ToolTipText     =   "Download Selected Video"
            Top             =   0
            Width           =   2415
         End
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
         Name            =   "Verdana"
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
      Top             =   2280
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
         MouseIcon       =   "Form1.frx":6094
      End
   End
   Begin Project1.Tray Tray 
      Left            =   0
      Top             =   2520
      _extentx        =   847
      _extenty        =   847
      pictureicon     =   "Form1.frx":61F6
   End
   Begin VB.Timer trmGetClipData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3000
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
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8160
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   34
         Top             =   1320
         Width           =   735
         Begin VB.CommandButton cmdVisit 
            Caption         =   "Visit"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   35
            ToolTipText     =   "Visit YouTube Channel"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtLength 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         TabIndex        =   24
         ToolTipText     =   "Video's Duration"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtView 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
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
         Enabled         =   0   'False
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
         Width           =   2655
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
         Left            =   3360
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Hit Enter When You Are Done"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
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
         Name            =   "Verdana"
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
      Top             =   2040
      Width           =   10695
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   1095
         ScaleWidth      =   10215
         TabIndex        =   36
         Top             =   3240
         Width           =   10215
         Begin VB.CommandButton cmdDown 
            Caption         =   "Download Selected Video"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   39
            ToolTipText     =   "Download The Selected Video"
            Top             =   480
            Width           =   7095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Open Selected Video In Browser"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   38
            ToolTipText     =   "Open The Video In Your Default Browser"
            Top             =   0
            Width           =   7095
         End
         Begin VB.CommandButton cmdAddDown 
            Caption         =   "Add To Download List >>>"
            Enabled         =   0   'False
            Height          =   975
            Left            =   7080
            TabIndex        =   37
            ToolTipText     =   "Add Selected Video To Download List "
            Top             =   0
            Width           =   3135
         End
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
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
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '获取鼠标在屏幕上的坐标
Private Declare Function ClientToScreen Lib "user32 " (ByVal hWnd As Long, lpPoint As POINTAPI) As Long '用户坐标转换为屏幕坐标
Private Type POINTAPI
    x   As Long
    y   As Long
End Type
Private Type RECT
    Left   As Long
    Top   As Long
    Right   As Long
    Bottom   As Long
End Type

Dim Pm As POINTAPI                                                              'Mouse Cursor Coordinate
Dim PRct As RECT
Dim OldPicSize As POINTAPI                                                      'Old Picture2 Size
Dim OldFrmSize As POINTAPI                                                      ' Old Form Size
Dim YoutubeDLink As String, WebSwf As String, DecodeString As String, WebHtml As String
Public DlFileEx As String
Dim URLFileSize As String                                                       'For Getting URL size True=Get
'Dim URLCheckTmp As Boolean
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long

Dim InetFail As Boolean
'Dim PrevWndProc_frmmain As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


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
        Dim i As Long
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Selected = True Then
                .ListItems.Remove i
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
    Dim i As Long
    Load frmDownload
    Dim x As ListItem
    For i = 1 To lvwDownloadList.ListItems.Count
        'Debug.Print lvwDownloadList.ListItems.Item(i)
        'Debug.Print lvwDownloadList.ListItems.Item(i).SubItems(1)
        Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.ListItems.Item(i))
        x.SubItems(1) = lvwDownloadList.ListItems.Item(i).SubItems(1)
        x.SubItems(2) = lvwDownloadList.ListItems.Item(i).SubItems(2)
        x.SubItems(3) = lvwDownloadList.ListItems.Item(i).SubItems(3)
        x.SubItems(4) = lvwDownloadList.ListItems.Item(i).SubItems(4)
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

Function ExtractMatch(Text, Pattern)
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

Private Sub Form_Activate()
    Static Rn As Integer
    If Rn = 0 Then
        Rn = 1
    Else
        Exit Sub
    End If
    If CheckUpdate = True And mnuCheckUpAuto.Checked = True Then
        If MsgBox("New Version Found!" & vbCrLf & "Update Now?", vbInformation + vbYesNo, App.Title) = vbYes Then
            frmUpdate.Show 1
        End If
    End If
End Sub


Private Sub form_Initialize()
    InitCommonControls
    
End Sub

Private Sub Form_Load()
    Inet1.RequestTimeout = 30
    'cmbDownloadOption.AddItem "Please select..."
    If App.LogMode = 0 Then
        MsgBox "Don't Run At IDE Please!", vbInformation
        RemoveMenu GetSystemMenu(hWnd, 0), SC_CLOSE, MF_REMOVE
        'End
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
    
    With OldPicSize
        .x = Picture2.Width
        .y = Picture2.Height
    End With
    With OldFrmSize
        .x = Me.Width
        .y = Me.Height
    End With
    
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
    
    'Unload Form2
    'Unload frmConvert
    Unload frmDownload
    Unload frmMain
    'Unload frmPlay
    Unload frmAbout
    ' Unload frmSupportedFileType
    Unload frmUpdate
    Unload Me
    Call TerminateProcess(GetCurrentProcess, ByVal 0&)
    
End Sub








Private Sub Inet1_StateChanged(ByVal State As Integer)
    'On Error Resume Next
    'Dim vtData As Variant
    
    Select Case State
    Case icError
        lstAdd "Error:" & Inet1.ResponseInfo & "(Error No. " & Inet1.ResponseCode & ")"
        
        Inet1.Cancel
        txtLink.Text = ""
        txtID.Text = ""
        'txtLike.Text = ""
        'txtDislikes.Text = ""
        txtTitle.Text = ""
        txtUploader.Text = ""
        txtHome.Text = ""
        Picture2.Picture = LoadPicture("")
        Dim x As Integer
        'For x = 0 To 3
        'txtDownload(x).Text = ""
        'optDownload(x).Enabled = False
        'Next
        
        Exit Sub
        'Case icResponseCompleted
        'Dim strData As String
        'Dim bDone As Boolean
        
        'bDone = False
        
        'vtData = Inet1.GetChunk(1024, icString)
        'DoEvents
        
        'Do While Not bDone
        'strData = strData & vtData
        
        'vtData = Inet1.GetChunk(1024, icString)
        'DoEvents
        
        'If Len(vtData) = 0 Then
        'bDone = True
        'End If
        'lstadd "Getting Web Soure Code"
        'lstState.ListIndex = lstState.NewIndex
        'Loop
        'WebHtml = strData
        'lstadd "Done Getting Web Source Code"
        'lstState.ListIndex = lstState.NewIndex
        
    End Select
    
End Sub

Public Function LoadPicture(ByVal strFileName As String) As Picture
    
    Dim IID  As TGUID
    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    On Error GoTo ERR_LINE
    OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
    Exit Function
ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)
End Function


' Copy the image from pic_src into pic_dst so
' it fits and has the same aspect ratio as the
' original picture.
Public Sub FitPictureToBox(ByVal pic_src As PictureBox, pic_dst As PictureBox)
    On Error Resume Next
    Dim aspect_src As Single
    Dim wid As Single
    Dim hgt As Single
    
    ' get the original picture's aspect ratio.
    aspect_src = pic_src.ScaleWidth / pic_src.ScaleHeight
    
    ' get the size available.
    wid = pic_dst.ScaleWidth
    hgt = pic_dst.ScaleHeight
    
    ' Adjust the wid/hgt ratio to match aspect_src.
    If wid / hgt > aspect_src Then
        ' The area is too short and wide.
        ' Make it narrower.
        wid = aspect_src * hgt
    Else
        ' The area is too tall and thin.
        ' Make it shorter.
        hgt = wid / aspect_src
    End If
    
    ' Center the image at the correct size.
    pic_dst.Cls
    pic_dst.PaintPicture pic_src.Picture, _
    (pic_dst.ScaleWidth - wid) / 2, _
    (pic_dst.ScaleHeight - hgt) / 2, _
    wid, hgt
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


Private Sub InetUpdate_StateChanged(ByVal State As Integer)
    If State = icError Then
        InetFail = True
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
    Dim i As Long
    For i = 1 To tvwQuality.Nodes.Count
        tvwQuality.Nodes(i).Expanded = False
    Next i
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
        Dim i As Long
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Selected = True Then
                .ListItems.Remove i
            End If
        Next
    End With
    cmdDownAll.Enabled = False
    cmdDownSelected.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub mnuDownloadAll_Click()
    Dim i As Long
    Load frmDownload
    Dim x As ListItem
    For i = 1 To lvwDownloadList.ListItems.Count
        'Debug.Print lvwDownloadList.ListItems.Item(i)
        'Debug.Print lvwDownloadList.ListItems.Item(i).SubItems(1)
        Set x = frmDownload.lvwDownload.ListItems.Add(, , lvwDownloadList.ListItems.Item(i))
        x.SubItems(1) = lvwDownloadList.ListItems.Item(i).SubItems(1)
        x.SubItems(2) = lvwDownloadList.ListItems.Item(i).SubItems(2)
        x.SubItems(3) = lvwDownloadList.ListItems.Item(i).SubItems(3)
        x.SubItems(4) = lvwDownloadList.ListItems.Item(i).SubItems(4)
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
    Dim i As Long
    For i = 1 To tvwQuality.Nodes.Count
        tvwQuality.Nodes(i).Expanded = True
    Next i
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
    If InStr(LCase$(Clipboard.GetText), "youtube") <> 0 And InStr(LCase$(Clipboard.GetText), "http://") And LCase$(txtLink.Text) <> LCase$(Clipboard.GetText) Then
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
    Dim i As Long
    If tvwQuality.SelectedItem.Children <> 0 Then Exit Sub
    'If tvwQuality.SelectedItem.Child.Selected = False Then Exit Sub
    'If tvwQuality.SelectedItem.Child.Parent.Selected Then Exit Sub
    
    If tvwQuality.SelectedItem.Children = 0 Then
        
        For i = 1 To tvwQuality.Nodes.Count
            If tvwQuality.Nodes(i).Selected Then
                lstAdd "Quality Selected. Analyzing"
                
                txtDownloadLink.Text = tvwQuality.Nodes(i).Text
                'Dim h As Long
                'For h = 0 To cmbDownloadOption.ListCount - 1
                'If cmbDownloadOption.List(h) = tvwQuality.Nodes(i).Text Then
                'cmbDownloadOption.ListIndex = h
                'Exit For
                'End If
                '    Next
                'cmbDownloadOption.List(cmbDownloadOption.ListIndex) = tvwQuality.Nodes(i).Text
            End If
        Next i
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

Private Sub txtHome_Change()
    If Len(txtHome.Text) <> 0 And txtHome.Text <> "Failed." Then
        cmdVisit.Enabled = True
    Else
        cmdVisit.Enabled = False
    End If
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    If txtLink.Text <> "" And KeyCode = 13 Then
        txtLink.Enabled = False
        tvwQuality.SetFocus
        Dim x As Integer
        CleanUp
        If InStr(LCase(txtLink.Text), "youtube") = 0 Then lstAdd "Failed To Get Video Info.": txtLink.Text = "": txtLink.SetFocus: Exit Sub
        lstAdd "Getting Video's ID"
        
        Dim VideoId  As String
        
        txtID.Text = ExtractMatch(txtLink.Text, "v=([A-Za-z0-9-_]+)")
        
        VideoId = ExtractMatch(txtLink.Text, "v=([A-Za-z0-9-_]+)")
        
        If txtID.Text = "" Then txtID.Text = "Failed.": lstAdd "Failed To Get Video Info."
        If txtID.Text = "Failed." Then Exit Sub
        
        
        lstAdd "Getting Web Source Code"
        
        If txtID.Text <> "" Then
            Dim WebBin() As Byte
            Inet1.Cancel
            WebBin = Inet1.OpenURL("http://youtube.com/get_video_info?video_id=" & VideoId, icByteArray)
            Do Until UBound(WebBin) > 0
                DoEvents
            Loop
            WebHtml = Utf8ToUnicode(WebBin)
            'Inet1.Cancel
            'Dim WebBin() As Byte
            'Inet1.Execute txtLink.Text, "Get"
            'WebBin = Inet1.OpenURL(txtLink.Text, icByteArray)
            'Do Until UBound(WebBin) > 0
            'DoEvents
            'Loop
            'Debug.Print WebBin
            'WebHtml = Utf8ToUnicode(WebBin)
            
        End If
        If InStr(WebHtml, "status=fail&errorcode=") Then
            lstAdd "Sorry,This Video Contain Copyrighted Information, Cannot Be Downloaded"
            CleanUp
            txtLink.Text = ""
            txtLink.Enabled = True
            Clipboard.Clear
            txtLink.SetFocus
            Exit Sub
        End If
        'lstAdd "Decoding Web Source Code"
        
        'Dim MaxWaitNum As Long
        
        'Do Until WebHtml <> ""                                                  'lstState.List(lstState.ListCount - 1) = "Done Getting Web Source Code"
        'If MaxWaitNum > 10000 Then Exit Sub
        'MaxWaitNum = MaxWaitNum + 1
        'DoEvents
        'Loop                                                                        '
        
        ' get rid of all newlines (vbscript regex engine doesn't like them)
        WebHtml = Replace(WebHtml, Chr(9), "")
        WebHtml = Replace(WebHtml, vbCrLf, "")
        WebHtml = Replace(WebHtml, vbLf, "")
        WebHtml = Replace(WebHtml, vbCr, "")
        
        
        'Dim x As Long
        For x = 0 To 9
            WebHtml = URLDecode(WebHtml)
        Next
        
        
        Dim VideoTitle As String
        VideoTitle = Trim$(Mid$(WebHtml, InStrRev(WebHtml, "&title=") + Len("&title="), Len(WebHtml) - InStrRev(WebHtml, "&title=") + 1)) 'Mid$(WebHtml, InStr(WebHtml, "<title>") + 7, InStr(WebHtml, "- YouTube  </title>") - InStr(WebHtml, "<title>") - 7))
        txtTitle.Text = VideoTitle                                              'Mid$(WebHtml, InStrRev(WebHtml, "&title=") + Len("&title="), Len(WebHtml) - InStrRev(WebHtml, "&title=") + 1) 'Trim$(Mid$(WebHtml, InStr(WebHtml, "<title>") + 7, InStr(WebHtml, "- YouTube  </title>") - InStr(WebHtml, "<title>") - 7))
        
        If txtTitle.Text Like "*[/\?:*|""<>]*" Then
            Dim ReX(8) As String
            ReX(0) = "/"
            ReX(1) = "\"
            ReX(2) = "?"
            ReX(3) = ":"
            ReX(4) = "*"
            ReX(5) = "|"
            ReX(6) = """"
            ReX(7) = "<"
            ReX(8) = ">"
            For x = LBound(ReX) To UBound(ReX)
                txtTitle.Text = Replace(txtTitle.Text, ReX(x), " - ")
            Next
        End If
        'Debug.Print WebHtml
        lstAdd "Getting Video Length"
        
        'txtLength.Text = Val(Mid$(WebHtml, InStr(WebHtml, """length_seconds"":") + Len("""length_seconds"": "))) & " Seconds"
        
        Dim VideoLength As String
        VideoLength = ConvertToTime(Val(ExtractMatch(WebHtml, "&length_seconds=([0-9]+)"))) 'ConvertToTime(Val(Mid$(WebHtml, InStr(WebHtml, """length_seconds"":") + Len("""length_seconds"": "))))
        txtLength.Text = VideoLength                                            'ConvertToTime(Val(ExtractMatch(WebHtml, "&length_seconds=([0-9]+)"))) 'ConvertToTime(Val(Mid$(WebHtml, InStr(WebHtml, """length_seconds"":") + Len("""length_seconds"": "))))
        If txtTitle.Text = "" Then txtTitle.Text = "Failed."
        
        lstAdd "Getting Video ScreenShot"
        'LoadWebImage Me.Picture1, Me.Picture2
        Set Picture1.Picture = LoadPicture(Mid$(WebHtml, InStrRev(WebHtml, "=", InStrRev(WebHtml, ".jpg")) + Len("="), InStrRev(WebHtml, ".jpg") - InStrRev(WebHtml, "=", InStrRev(WebHtml, ".jpg")) + Len(".jpg") - 1))
        'Set Picture1.Picture = LoadPicture(Mid$(WebHtml, InStr(WebHtml, "<meta property=""og:image"" content=") + 35, InStr(WebHtml, "mqdefault.jpg"">") - InStr(WebHtml, "<meta property=""og:image"" content=") - 35 + 13))
        'Debug.Print Mid$(WebHtml, InStr(WebHtml, "<meta property=""og:image"" content=") + 35, InStr(WebHtml, "mqdefault.jpg"">") - InStr(WebHtml, "<meta property=""og:image"" content=") - 35 + 13)
        FitPictureToBox Picture1, Picture2
        
        
        'lstAdd "Getting Video's Likes and Dislikes Numbers"
        'txtLike.Text = Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""likes"">") + 20, InStr(WebHtml, "<span class=""dislikes"">") - InStr(WebHtml, "<span class=""likes"">") - 20), ",", ""))
        'txtLike.Text = ,  - InStr(WebHtml, "<span class=""likes"">")) 'Mid$(WebHtml, InStr(WebHtml, "<span class=""watch-likes-dislikes""><span class=""likes"">") + 55, InStr(WebHtml, "</span> likes, <span class=""dislikes"">") - InStr(WebHtml, "<span class=""watch-likes-dislikes""><span class=""likes"">") - 55)
        'If txtLike.Text = "" Then txtLike.Text = "Failed."
        
        'Dim VideoLikes As String
        'VideoLikes = Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""likes"">") + 20, InStr(WebHtml, "<span class=""dislikes"">") - InStr(WebHtml, "<span class=""likes"">") - 20), ",", ""))
        'txtDislikes.Text = Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""dislikes"">") + 23), ",", ""))
        'Mid$(WebHtml, InStr(WebHtml, "<span class=""dislikes"">") + 23, InStr(WebHtml, "</span> dislikes") - InStr(WebHtml, "<span class=""dislikes"">") - 23)
        'If txtDislikes.Text = "" Then txtDislikes.Text = "Failed."
        
        'Dim VideoDislikes As String
        'VideoDislikes = Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""dislikes"">") + 23), ",", ""))
        
        lstAdd "Getting Uploader's Username"
        
        'Dim usertmp As String
        'usertmp = Mid$(WebHtml, InStr(WebHtml, "yt-uix-button-group"))
        'txtUploader.Text = ExtractMatch(WebHtml, """ptchn"": ""([]0-9a-zA-z~!@#\$%\^&\*\(\)\-\+=~`\{\};'/\.,\?\<\>\|]+)""")
        'If txtUploader.Text = "" Then
        'txtUploader.Text = usertmp
        'txtUploader.Text = Mid$(txtUploader.Text, InStr(txtUploader.Text, "button href=") + Len("button href="""), InStr(txtUploader.Text, """ type=""button""") - Len(""" type=""button"" class=""") - Len("button href="""))
        'txtUploader.Text = ExtractMatch(txtUploader.Text, "user\/([A-Za-z0-9-_]+)")
        
        Dim VideoUploader As String
        VideoUploader = Mid$(WebHtml, InStr(WebHtml, "&author=") + Len("&author="), InStr(WebHtml, "&muted=") - InStr(WebHtml, "&author=") - Len("&author=")) 'ExtractMatch((Mid$(usertmp, InStr(usertmp, "button href=") + Len("button href="""), InStr(usertmp, """ type=""button""") - Len(""" type=""button"" class=""") - Len("button href="""))), "user\/([A-Za-z0-9-_]+)")
        txtUploader.Text = VideoUploader
        'Debug.Print VideoUploader
        'End If
        'Mid$(WebHtml, InStr(WebHtml, "'VIDEO_USERNAME': """) + 19, InStr(WebHtml, """    });    yt.net.ajax.setToken('") - InStr(WebHtml, "'VIDEO_USERNAME': """) - 19)
        
        ', InStr(WebHtml, "yt.net.ajax.setToken") - InStr(WebHtml, "'VIDEO_USERNAME': """)) 'Mid$(WebHtml, InStr(WebHtml, "class=""yt-user-name author"" rel=""author""  dir=""ltr"">") + 52, InStr(WebHtml, "</a> on <span id=""eow-date""") - InStr(WebHtml, "class=""yt-user-name author"" rel=""author""  dir=""ltr"">") - 52)
        If Len(txtUploader.Text) = 0 Then txtUploader.Text = "Failed."
        'Debug.Print Len(Trim(txtUploader.Text))
        lstAdd "Getting Uploader Homepage"
        
        If txtUploader.Text <> "Failed." Then
            txtHome.Text = "http://www.youtube.com" & "/user/" & txtUploader.Text
            'Dim VideoHomepage As String
            'VideoHomepage = "http://www.youtube.com/user/" & VideoUploader
        Else
            txtHome.Text = "Failed."
        End If
        'Mid$(WebHtml, InStr(WebHtml, "<span class=""author """) + 41, InStr(WebHtml, "class=""yt-user-name """) - InStr(WebHtml, "<span class=""author """) - 43)
        'Mid$(WebHtml, InStr(WebHtml, "<a id=""watch-userbanner"" rel=""author"" href="""), InStr(WebHtml, """ title=" & IIf(Len(txtUploader.Text) <> 0, txtUploader.Text, "") & """><strong>") - InStr(WebHtml, "<a id=""watch-userbanner"" rel=""author"" href="""))
        'InetCheckedLink.Cancel
        'InetCheckedLink.Execute txtHome.Text, "GET"
        
        'Do Until URLCheckTmp <> ""
        'DoEvents
        'Loop
        
        'Debug.Print VideoHomepage
        lstAdd "Getting Video Views"
        Dim VideoViews As String
        VideoViews = ExtractMatch(WebHtml, "&view_count=([0-9]+)")              'Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""watch-view-count"">      <strong>") + Len("<span class=""watch-view-count"">      <strong>")), ",", ""))
        txtView.Text = VideoViews                                               'Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""watch-view-count"">      <strong>") + Len("<span class=""watch-view-count"">      <strong>")), ",", ""))
        
        
        'txtInfo = "[Video ID]: " & VideoId & vbCrLf & "[VideoTitle]: " & VideoTitle & vbCrLf & "[VideoLength]: " & _
        VideoLength & vbCrLf & "[Likes]: " & VideoLikes & vbCrLf & "[Video Dislikes]: " & VideoDislikes & _
        vbCrLf & "[Video Uploader]: " & VideoUploader & vbCrLf & "[Uploader's Channel]: " & VideoHomepage & vbCrLf & "[Views]: " & VideoViews
        
        'Debug.Print Mid$(WebHtml, InStr(WebHtml, "<span class=""watch-view-count"">      <strong>"))
        '= Val(Replace(Mid$(WebHtml, InStr(WebHtml, "<span class=""watch-view-count"">") + 43), ",", ""))
        lstAdd "Getting Video Download Link"
        'DecodeString = Mid$(WebHtml, InStr(WebHtml, "swf"), InStr(WebHtml, "document.getElementById('watch-player').innerHTML = swf") - InStr(WebHtml, "swf") + Len("document.getElementById('watch-player').innerHTML = swf"))
        'lstAdd "Decoding Video Download Link"
        'For x = 0 To 10
        'DecodeString = URLUtility.URLDecode(DecodeString, True)
        'Next
        
        'Debug.Print DecodeString
        'Dim DLURL() As String
        'DLURL = Split(Mid$(DecodeString, InStr(DecodeString, "url_encoded_fmt_stream_map"), InStr(DecodeString, "document.getElementById('watch-player').innerHTML = swf") - InStr(DecodeString, "url_encoded_fmt_stream_map")), "url", , vbTextCompare) 'Split(Mid$(DecodeString, InStr(DecodeString, "url_encoded_fmt_stream_map"), InStr(DecodeString, "watermark") - InStr(DecodeString, "url_encoded_fmt_stream_map") + 14), "url", , vbTextCompare)
        
        'MsgBox InStr(DecodeString, "url_encoded_fmt_stream_map")
        'YoutubeDLink = Right(DLURL(UBound(DLURL)), Len(DLURL(UBound(DLURL))) - 1)
        lstAdd "Splitting Video Download Link"
        
        'Dim YoutubeFDL() As String
        
        'MsgBox DLURL(1)
        'ReDim YoutubeFDL(UBound(DLURL) - 1)
        'For x = LBound(DLURL) + 2 To UBound(DLURL)
        'If InStr(DLURL(x), "o-o.preferred") <> 0 Then
        
        'If x = UBound(DLURL) Then
        'cmbDownloadOption.AddItem Mid$(DLURL(UBound(DLURL)), 2, InStr(Mid$(DLURL(UBound(DLURL)), InStr(DLURL(UBound(DLURL)), "type=video")), "&") + InStr(DLURL(UBound(DLURL)), "type=video") - 3)
        'Debug.Print Mid$(DLURL(UBound(DLURL)), 2, InStr(Mid$(DLURL(UBound(DLURL)), InStr(DLURL(UBound(DLURL)), "type=video")), "&") + InStr(DLURL(UBound(DLURL)), "type=video") - 3)
        'YoutubeFDL(x) = Mid$(DLURL(UBound(DLURL)), 2, InStr(Mid$(DLURL(UBound(DLURL)), InStr(DLURL(UBound(DLURL)), "type=video")), "&") + InStr(DLURL(UBound(DLURL)), "type=video") - 3)
        ' ElseIf InStrRev(DLURL(x), "; ") <> 0 Then
        'cmbDownloadOption.AddItem Mid$(DLURL(x), 2, InStrRev(DLURL(x), ";") - 2)
        'Debug.Print Mid$(DLURL(x), 2, InStrRev(DLURL(x), ";") - 2)
        'YoutubeFDL(x) = Mid$(DLURL(x), 2, InStrRev(DLURL(x), ";") - 2)
        'ElseIf InStrRev(DLURL(x), "&") <> 0 Then
        'cmbDownloadOption.AddItem Mid$(DLURL(x), 2, InStrRev(DLURL(x), "&") - 2)
        'Debug.Print Mid$(DLURL(x), 2, InStrRev(DLURL(x), "&") - 2)
        'YoutubeFDL(x) = Mid$(DLURL(x), 2, InStrRev(DLURL(x), "&") - 2)
        'End If
        
        'End If
        'Next
        Dim YoutubeFDL() As String
        ReDim YoutubeFDL(0)
        Do Until InStr(WebHtml, ";") = 0
            If InStr(Mid$(WebHtml, InStr(WebHtml, "url="), InStr(WebHtml, ";") - InStr(WebHtml, "url=")), "video/x-flv&itag=") Then
                'Combo1.AddItem Mid$(WebHtml, InStr(WebHtml, "url="), InStr(WebHtml, "video/x-flv&itag=") - InStr(WebHtml, "url=") + Len("video/x-flv"))
                YoutubeFDL(UBound(YoutubeFDL)) = Mid$(WebHtml, InStr(WebHtml, "url=") + Len("url="), InStr(WebHtml, "video/x-flv&itag=") - InStr(WebHtml, "url=") + Len("video/x-flv") - Len("url="))
                WebHtml = Replace(WebHtml, Mid$(WebHtml, InStr(WebHtml, "url="), InStr(WebHtml, "video/x-flv&itag=") - InStr(WebHtml, "url=") + Len("video/x-flv")), "")
                
            Else
                ' Combo1.AddItem Mid$(WebHtml, InStr(WebHtml, "url="), InStr(WebHtml, ";") - InStr(WebHtml, "url="))
                YoutubeFDL(UBound(YoutubeFDL)) = Mid$(WebHtml, InStr(WebHtml, "url=") + Len("url="), InStr(WebHtml, ";") - InStr(WebHtml, "url=") - Len("url="))
                WebHtml = Replace(WebHtml, Mid$(WebHtml, InStr(WebHtml, "url="), InStr(WebHtml, ";") - InStr(WebHtml, "url=") + 1), "")
                
            End If
            ReDim Preserve YoutubeFDL(UBound(YoutubeFDL) + 1)
            DoEvents
            
        Loop
        
        
        Dim TVWNode As Node
        
        lstAdd "Analyzing And Catergorizing Video Quality"
        
        Dim tmp As String
        For x = 1 To UBound(YoutubeFDL) - 1                                     'cmbDownloadOption.ListCount - 1
            If YoutubeFDL(x) <> "" Then
                tmp = LCase(ExtractMatch(YoutubeFDL(x), "&quality=([a-zA-Z0-9]*)"))
                If tmp <> "" Then
                    Set TVWNode = tvwQuality.Nodes.Add(, , tmp, tmp)
                    Set TVWNode = tvwQuality.Nodes.Add(tmp, tvwChild, , YoutubeFDL(x))
                End If
                'Debug.Print cmbDownloadOption.List(x)
                'Select Case LCase(ExtractMatch(cmbDownloadOption.List(X), "&quality=([a-zA-Z0-9]*)"))
                
                'Case "hd720"
                'Set TVWNode = tvwQuality.Nodes.Add(, , "hd720", "hd720")
                'Set TVWNode = tvwQuality.Nodes.Add("hd720", tvwChild, , cmbDownloadOption.List(X))
                'Case "large"
                'Set TVWNode = tvwQuality.Nodes.Add(, , "large", "large")
                'Set TVWNode = tvwQuality.Nodes.Add("large", tvwChild, , cmbDownloadOption.List(X))
                'Case "medium"
                'Set TVWNode = tvwQuality.Nodes.Add(, , "medium", "medium")
                'Set TVWNode = tvwQuality.Nodes.Add("medium", tvwChild, , cmbDownloadOption.List(X))
                'Case "small"
                'Set TVWNode = tvwQuality.Nodes.Add(, , "small", "small")
                'Set TVWNode = tvwQuality.Nodes.Add("small", tvwChild, , cmbDownloadOption.List(X))
                'End Select
                DoEvents
            End If
        Next
        
        'Debug.Print DLURL(UBound(DLURL))
        'cmbDownloadOption.RemoveItem cmbDownloadOption.ListCount - 1
        'cmbDownloadOption.AddItem Mid$(DLURL(UBound(DLURL)), 2, InStr(Mid$(DLURL(UBound(DLURL)), InStr(DLURL(UBound(DLURL)), "type=video")), "&") + InStr(DLURL(UBound(DLURL)), "type=video") - 3)
        'MsgBox cmbDownloadOption.List(8)
        'MsgBox cmbDownloadOption.List(UBound(DLURL) - 2)
        'lstAdd (cmbDownloadOption.ListCount - 1) & " Download Link(s) Found."
        lstAdd UBound(YoutubeFDL) - 1 & " Download Link(s) Found."
        'lstadd "Formatting Video Download Link."
        'Debug.Print InStr(cmbDownloadOption.List(cmbDownloadOption.ListCount - 1), "&")
        
        'cmbDownloadOption.AddItem Mid$(cmbDownloadOption.List(cmbDownloadOption.ListCount - 1), 1, InStr(cmbDownloadOption.List(cmbDownloadOption.ListCount - 1), "type=video/x-flv") + 15)
        
        
        'For x = 0 To 7
        'Debug.Print cmbDownloadOption.List(x)
        'Next
        
        'Debug.Print UBound(DLURL) & cmbDownloadOption.List(UBound(DLURL))
        
        'txtDownload(0).Text = Mid(txtDownload(0).Text, 1, InStrRev(txtDownload(0), ";") - 1)
        'txtDownload(1).Text = Mid(txtDownload(1).Text, 1, InStrRev(txtDownload(1), "&") - 1)
        
        'For x = 0 To 2
        'If txtDownload(x).Text <> "" Then
        'If mnuNoVideoSize.Checked = False Then
        'lstadd "Getting Video File Size"
        'InetFileSize.Cancel
        'InetFileSize.Execute txtDownload(x).Text, "GET"
        'Do Until URLFileSize <> ""
        'DoEvents
        'Loop
        'optDownload(x).Caption = "Download this" & vbCrLf & URLFileSize
        'URLFileSize = ""
        'End If
        'optDownload(x).Enabled = True
        'Else
        'txtDownload(x).Text = "Failed."
        'End If
        'Next
        'txtDownload.Text = YoutubeDLink
        'If txtDownload(0).Text = "" Then txtDownload(0).Text = "Failed.": cmdDown.Enabled = False
        'cmbDownloadOption.Enabled = True
        lstAdd "Everything Done"
        mnuExpand_Click
        'cmbDownloadOption.SetFocus
        txtLink.Enabled = True
    Else
        txtLink.SetFocus
    End If
    Exit Sub
goterr:
    lstAdd "Program Error!"
    CleanUp
    txtLink.Enabled = True
End Sub

Function VisitURL(VisitUrlink As String)
    On Error Resume Next
    Shell "rundll32.exe url.dll,FileProtocolHandler " & Chr(34) & VisitUrlink & Chr(34), 1
End Function


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

Private Function CheckUpdate() As Boolean
    
    Dim iniF() As Byte
    iniF = InetUpdate.OpenURL("http://dl.dropbox.com/u/43619472/%E6%89%B9%E5%A4%84%E7%90%86/VB6/Youtube%20Downloader/Updates/Update.ini", icByteArray)
    If InetFail = True Then Exit Function
    SaveToFile Environ$("temp") & "\AGYTVGU.ini", iniF
    
    
    
    Dim x As Double
    Dim NewVer As String, tmpVer As String
    NewVer = GetIni("Version", "Ver", App.Major & "." & App.Minor & "." & App.Revision, Environ$("temp") & "\AGYTVGU.ini")
    
    tmpVer = Replace(NewVer, ".", "")
    If Val(tmpVer) > (App.Major & App.Minor & App.Revision) Then
        CheckUpdate = True
    End If
End Function

Private Sub lstAdd(ByVal InName As String)
    lstState.AddItem Date & "-" & Time & ": " & InName
    lstState.ListIndex = lstState.NewIndex
    
End Sub

