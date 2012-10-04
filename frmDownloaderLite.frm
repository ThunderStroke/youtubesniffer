VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDownloaderLite 
   Caption         =   "AG YouTube Video Grabber - Lite - Downloader"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "frmDownloaderLite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScrollCon 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   10035
      TabIndex        =   1
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
         TabIndex        =   2
         Top             =   0
         Width           =   9375
      End
   End
   Begin VB.Timer tmrScrolling 
      Interval        =   10
      Left            =   120
      Top             =   2160
   End
   Begin MSComctlLib.ListView lvwDonwloadList 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "frmDownloaderLite.frx":5F32
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDownloaderLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
