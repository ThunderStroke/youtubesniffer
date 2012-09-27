VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainLite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG YouTube Video Grabber - Lite"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
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
   ScaleHeight     =   5265
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   5520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4680
      Width           =   5415
   End
   Begin MSComctlLib.ListView lvwDownloadLinks 
      Height          =   3135
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5530
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
         Object.Width           =   2540
      EndProperty
   End
   Begin InetCtlsObjects.Inet inetLinkInfo 
      Left            =   10920
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picScrollCon 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   10875
      TabIndex        =   2
      Top             =   600
      Width           =   10935
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
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10455
      End
   End
   Begin VB.Timer tmrScrolling 
      Enabled         =   0   'False
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
      Width           =   10935
   End
   Begin VB.PictureBox picScreenShot 
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
      Height          =   3135
      Left            =   240
      MouseIcon       =   "frmMainLite.frx":6094
      ScaleHeight     =   3075
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   1440
      Width           =   5055
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


Private Sub Form_Load()
    picScreenShot = LoadPicture("http://i4.ytimg.com/vi/oxqnFJ3lp5k/mqdefault.jpg")
    txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    lblTitle.AutoSize = True
    picScrollCon.BorderStyle = 0
End Sub

Private Sub inetLinkInfo_StateChanged(ByVal State As Integer)
    If State = 12 Then
        nVideoFileSize = VBStrFormatByteSize(inetLinkInfo.GetHeader("content-length"))
        strVideoType = inetLinkInfo.GetHeader("content-type")
    End If
    
End Sub

Private Sub lvwDownloadLinks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lvwItems As ListItem
    'If lvwDownloadLinks.SelectedItem.SubItems(2) <> "" Then
    lvwDownloadLinks.Enabled = False
    inetLinkInfo.Cancel
    inetLinkInfo.Execute lvwDownloadLinks.SelectedItem.SubItems(1), "GET"
    Do Until nVideoFileSize <> ""
        DoEvents
    Loop
    Text1.Text = Text1.Text & nVideoFileSize
    Text1.Text = Text1.Text & "*." & Replace(Replace(strVideoType, "x-", ""), "video/", "")
    lvwDownloadLinks.Enabled = True
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
        
        lblTitle.Caption = txtLink.Text
        SeperateSWF txtLink.Text
        lblTitle.Caption = LoadVideoTitle
        tmrScrolling.Enabled = True
        
        Dim strDownloadLinks() As String
        ProcessDownloadLinks strDownloadLinks
        Dim lvwItems As ListItem, nI As Long
        For nI = 0 To UBound(strDownloadLinks)
            Set lvwItems = lvwDownloadLinks.ListItems.Add(, , ExtractMatch(strDownloadLinks(nI), "&quality=([a-zA-Z0-9]*)"))
            lvwItems.SubItems(1) = strDownloadLinks(nI)
        Next
    End If
End Sub


Private Sub txtLink_LostFocus()
    If txtLink.Text = "" Then
        txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    End If
End Sub

Private Sub CleanUp()
    txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    lblTitle.Caption = ""
    tmrScrolling.Enabled = False
    picScreenShot.Picture = LoadPicture("")
    lvwDownloadLinks.ListItems.Clear
End Sub
