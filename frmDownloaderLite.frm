VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownloaderLite 
   Caption         =   "AG YouTube Video Grabber - Lite - Downloader"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloaderLite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   StartUpPosition =   2  'CenterScreen
   Begin AGYouTubeVideoGrabber.ctlProgressBar ctlProgressBar1 
      Height          =   150
      Left            =   120
      Top             =   3120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   265
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   9975
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   9960
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "..."
      Height          =   345
      Left            =   9600
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwDownloadList 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Video Title"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Video Format"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vidoe Quality"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Save To..."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Video Download Link"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtSaveTo 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9375
   End
   Begin VB.PictureBox picScrollCon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Label lblTitle 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   9375
      End
   End
   Begin VB.Timer tmrScrolling 
      Interval        =   10
      Left            =   10080
      Top             =   3480
   End
End
Attribute VB_Name = "frmDownloaderLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    On Error GoTo errCancel
    With cdSave
        .FileName = txtSaveTo.Text
        .Flags = 2
        .CancelError = True
        .DefaultExt = "*." & frmMainLite.strExtension
        .Filter = frmMainLite.strExtension & " (" & frmMainLite.strExtension & ")" & "|" & "*." & frmMainLite.strExtension
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        txtSaveTo.Text = .FileName
    End With
errCancel:
    If Err.Number = cdlCancel Then
        txtSaveTo.Text = App.Path & "\" & frmMainLite.strVideoTitle & frmMainLite.strExtension
    End If
End Sub

Private Sub Form_Load()
    lblTitle.AutoSize = True
    picScrollCon.BorderStyle = 0
    
    lblTitle.Caption = frmMainLite.strVideoTitle
    
End Sub

Private Sub tmrScrolling_Timer()
    If picScrollCon.Left - lblTitle.Left < lblTitle.Width Then
        lblTitle.Left = lblTitle.Left - 15
    Else
        lblTitle.Left = picScrollCon.Left + picScrollCon.ScaleWidth + 100
    End If
End Sub


