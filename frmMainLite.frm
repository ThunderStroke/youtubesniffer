VERSION 5.00
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
      Height          =   3135
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   5055
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
      Left            =   5520
      Top             =   1320
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
      Left            =   6000
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

Private Sub Form_Load()
    picScreenShot = LoadPicture("http://i4.ytimg.com/vi/oxqnFJ3lp5k/mqdefault.jpg")
    txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    lblTitle.AutoSize = True
    picScrollCon.BorderStyle = 0
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
        tmrScrolling.Enabled = True
        lblTitle.Caption = txtLink.Text
    End If
End Sub


Private Sub txtLink_LostFocus()
    If txtLink.Text = "" Then
        txtTips txtLink, "Paste the YouTube video link here... Press enter when you're done", True
    End If
End Sub

