VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainLite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG YouTube Video Grabber - Lite"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainLite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   4560
      Width           =   5295
   End
   Begin VB.CommandButton cmdMoreInfo 
      Caption         =   "<<<More About This Video>>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Timer tmrGetClipData 
      Interval        =   100
      Left            =   720
      Top             =   1200
   End
   Begin MSComctlLib.ListView lvwDownloadLinks 
      Height          =   2775
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4895
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
      MouseIcon       =   "frmMainLite.frx":5F32
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Video Quality"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Video Download Link"
         Object.Width           =   5292
      EndProperty
   End
   Begin InetCtlsObjects.Inet inetLinkInfo 
      Left            =   9120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picScrollCon 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   10515
      TabIndex        =   2
      Top             =   600
      Width           =   10575
      Begin VB.Label lblTitle 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Timer tmrScrolling 
      Interval        =   10
      Left            =   240
      Top             =   1200
   End
   Begin VB.TextBox txtLink 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Text            =   "Paste the YouTube video link here..."
      Top             =   120
      Width           =   10575
   End
   Begin VB.PictureBox picScreenShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      MouseIcon       =   "frmMainLite.frx":6094
      ScaleHeight     =   2745
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
   End
End
Attribute VB_Name = "frmMainLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nVideoFileSize As String
Dim strVideoType As String
Public strVideoID As String, strVideoTitle As String, strVideoViews As String, strVideoLength As String
Public strVideoUploader As String, strVideoChannel As String, strDownloadLink As String


Private Sub cmdDownload_Click()
    strDownloadLink = lvwDownloadLinks.SelectedItem.SubItems(1)
    Load frmDownloaderLite
    frmDownloaderLite.Show
    
End Sub

Private Sub cmdMoreInfo_Click()
    Load frmMoreInfoLite
    frmMoreInfoLite.Show 1
End Sub

Private Sub Form_Load()
    
    txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    lblTitle.AutoSize = True
    picScrollCon.BorderStyle = 0
    lblTitle.Caption = Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
    Unload Me
End Sub

Private Sub inetLinkInfo_StateChanged(ByVal State As Integer)
    If State = 12 Then
        nVideoFileSize = VBStrFormatByteSize(inetLinkInfo.GetHeader("content-length"))
        strVideoType = Replace(Replace(inetLinkInfo.GetHeader("content-type"), "x-", ""), "video/", "")
    End If
    
End Sub


Private Sub lvwDownloadLinks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lvwItems As ListItem
    lblInfo.Caption = ""
    nVideoFileSize = ""
    strVideoType = ""
    'If lvwDownloadLinks.SelectedItem.SubItems(2) <> "" Then
    lvwDownloadLinks.Enabled = False
    inetLinkInfo.Cancel
    inetLinkInfo.Execute lvwDownloadLinks.SelectedItem.SubItems(1), "GET"
    Do Until nVideoFileSize <> ""
        DoEvents
    Loop
    
    lblInfo.Caption = lblInfo.Caption & nVideoFileSize & "      "
    lblInfo.Caption = lblInfo.Caption & "*." & strVideoType
    lblInfo.AutoSize = True
    lvwDownloadLinks.Enabled = True
    cmdDownload.Enabled = True
    cmdDownload.SetFocus
End Sub

Private Sub lvwDownloadLinks_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show
End Sub

Private Sub tmrGetClipData_Timer()
    On Error Resume Next
    If frmSettings.swhAutoClipLite.Value = False Then Exit Sub
    If ExtractMatch(Clipboard.GetText, "v=([A-Za-z0-9-_]+)") = "" Then Exit Sub
    If InStr(LCase$(Clipboard.GetText), "www.youtube.com") <> 0 And LCase$(txtLink.Text) <> LCase$(Clipboard.GetText) Then
        txtLink_GotFocus
        txtLink.Text = Clipboard.GetText
        If frmSettings.swhActAppLite.Value = True Then
            AppActivate Me.Caption
        End If
        txtLink_KeyDown 13, 0
    End If
    Exit Sub
End Sub


Private Sub tmrScrolling_Timer()
    If txtLink.Text = "Paste the YouTube video link here... Press enter when you're done" Then Exit Sub
    If picScrollCon.Left - lblTitle.Left < lblTitle.Width Then
        lblTitle.Left = lblTitle.Left - 15
    Else
        lblTitle.Left = picScrollCon.Left + picScrollCon.ScaleWidth + 100
    End If
    
End Sub

Private Sub txtLink_GotFocus()
    If txtLink.Text = "Paste the YouTube video link here... Press enter when you're done" Then
        txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", False
    End If
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Dim lvwItems As ListItem, nI As Long, strDownloadLinks() As String
        If txtLink.Text = "" Then CleanUp: Exit Sub
        txtLink.Enabled = False
        tmrGetClipData.Enabled = False
        picScreenShot.Picture = LoadPicture("")
        lvwDownloadLinks.ListItems.Clear
        lblTitle.Left = picScrollCon.Left + 150
        lvwDownloadLinks.ListItems.Clear
        lblInfo.Caption = ""
        lblTitle.Caption = txtLink.Text
        SeperateSWF txtLink.Text
        LoadPicScreenShot picScreenShot, picScreenShot
        LoadVideoInfoLite txtLink.Text, strVideoTitle, strVideoUploader, strVideoChannel, strVideoID, strVideoViews, strVideoLength
        lblTitle.Caption = strVideoTitle & "        " & IIf(strVideoUploader <> "", "By - " & strVideoUploader, "")
        ProcessDownloadLinks strDownloadLinks
        For nI = 0 To UBound(strDownloadLinks)
            Set lvwItems = lvwDownloadLinks.ListItems.Add(, , ExtractMatch(strDownloadLinks(nI), "&quality=([a-zA-Z0-9]*)"))
            lvwItems.SubItems(1) = strDownloadLinks(nI)
        Next
        txtLink.Enabled = True
        cmdMoreInfo.Enabled = True
        tmrGetClipData.Enabled = True
    End If
End Sub


Private Sub txtLink_LostFocus()
    If txtLink.Text = "" Then
        txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    End If
End Sub

Private Sub CleanUp()
    txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    lblTitle.Caption = Me.Caption
    picScreenShot.Picture = LoadPicture("")
    lvwDownloadLinks.ListItems.Clear
    lblInfo.Caption = ""
    cmdMoreInfo.Enabled = False
    cmdDownload.Enabled = False
End Sub
