VERSION 5.00
Begin VB.Form frmMoreInfoLite 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG YouTube Video Grabber - Lite - More Info"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMoreInfoLite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      BackColor       =   &H8000000F&
      Height          =   3375
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton cmdVisitURL 
      Caption         =   "Visit Author's Channel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   5055
   End
   Begin VB.PictureBox picScrollCon 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   10995
      TabIndex        =   1
      Top             =   120
      Width           =   11055
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
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   10215
      End
   End
   Begin VB.Timer tmrScrolling 
      Interval        =   10
      Left            =   720
      Top             =   1440
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
      Height          =   3015
      Left            =   120
      MouseIcon       =   "frmMoreInfoLite.frx":5F32
      ScaleHeight     =   2985
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   720
      Width           =   5055
   End
End
Attribute VB_Name = "frmMoreInfoLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVisitURL_Click()
    VisitURL frmMainLite.strVideoChannel
End Sub

Private Sub Form_Load()
    lblTitle.AutoSize = True
    picScreenShot.Picture = frmMainLite.picScreenShot.Picture
    lblTitle.Caption = frmMainLite.strVideoTitle
    picScrollCon.BorderStyle = 0
    LoadInfo
End Sub

Private Sub tmrScrolling_Timer()
    If picScrollCon.Left - lblTitle.Left < lblTitle.Width Then
        lblTitle.Left = lblTitle.Left - 15
    Else
        lblTitle.Left = picScrollCon.Left + picScrollCon.ScaleWidth + 100
    End If
End Sub

Private Sub LoadInfo()
    txtInfo.Text = "Video ID:                " & frmMainLite.strVideoID & vbCrLf
    txtInfo.Text = txtInfo.Text & "Video Author:           " & frmMainLite.strVideoUploader & vbCrLf
    txtInfo.Text = txtInfo.Text & "Video Channel:          " & frmMainLite.strVideoChannel & vbCrLf
    txtInfo.Text = txtInfo.Text & "Video Length:           " & frmMainLite.strVideoLength & vbCrLf
    txtInfo.Text = txtInfo.Text & "Video Views:            " & frmMainLite.strVideoViews
End Sub

