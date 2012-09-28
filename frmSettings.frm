VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AG YouTube Video Grabber - Settings"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLite 
      Height          =   4215
      Left            =   2760
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
      Begin AGYouTubeVideoGrabber.cltSwitch swhLiteDef 
         Height          =   345
         Left            =   4920
         TabIndex        =   27
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhAutoClipLite 
         Height          =   345
         Left            =   4920
         TabIndex        =   29
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhActAppLite 
         Height          =   345
         Left            =   4920
         TabIndex        =   31
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin VB.Label Label9 
         Caption         =   "Activate program when got link in clipboard"
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
         TabIndex        =   32
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "Auto get video link from clipboard"
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
         TabIndex        =   30
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   "Set lite version to default"
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
         TabIndex        =   28
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdLite 
      Caption         =   "Lite"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox picArrow 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox picTools 
      Height          =   4215
      Left            =   4080
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdUsage 
         Caption         =   "Optimize program's usage"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   5895
      End
      Begin VB.CommandButton cmdRestart 
         Caption         =   "Restart the program"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Tools"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdDownloader 
      Caption         =   "Downloader"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdater 
      Caption         =   "Updater"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "General"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Restore Default"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
   Begin VB.PictureBox picDownloader 
      Height          =   4215
      Left            =   2640
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   6255
      Begin AGYouTubeVideoGrabber.cltSwitch swhAutoOpenFolder 
         Height          =   345
         Left            =   4920
         TabIndex        =   14
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         value           =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Auto open folder after finish download"
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
         TabIndex        =   15
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.PictureBox picUpdater 
      Height          =   4215
      Left            =   2280
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdCheckUpdateNow 
         Caption         =   "Check For Updates Now"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   5895
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhAutoUpdate 
         Height          =   345
         Left            =   4920
         TabIndex        =   8
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhAutoOpenLink 
         Height          =   345
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin VB.Label Label2 
         Caption         =   "Auto open latest version download link"
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
         TabIndex        =   10
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Check updates when program start"
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
         TabIndex        =   9
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.PictureBox picGeneral 
      Height          =   4215
      Left            =   2880
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   -120
      Width           =   6255
      Begin AGYouTubeVideoGrabber.cltSwitch swhMinimize 
         Height          =   345
         Left            =   4920
         TabIndex        =   1
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhAutoGetLink 
         Height          =   345
         Left            =   4920
         TabIndex        =   3
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin AGYouTubeVideoGrabber.cltSwitch swhActivateApp 
         Height          =   345
         Left            =   4920
         TabIndex        =   5
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
      End
      Begin VB.Label Label4 
         Caption         =   "Activate program when got link in clipboard"
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
         TabIndex        =   6
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Auto video get link from clipboard"
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
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Minimize to tray"
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
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Line lnDivider 
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   4680
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckUpdateNow_Click()
    Unload frmUpdate
    Load frmUpdate
    frmUpdate.Show
End Sub

Private Sub cmdDefault_Click()
    swhMinimize.Value = True
    swhAutoGetLink.Value = True
    swhActivateApp.Value = True
    swhAutoUpdate.Value = True
    swhAutoOpenLink.Value = True
    swhAutoOpenFolder.Value = False
End Sub

Private Sub cmdDownloader_Click()
    picArrow.Top = cmdDownloader.Top
    picDownloader.Visible = True
    picGeneral.Visible = False
    picUpdater.Visible = False
    picTools.Visible = False
        picLite.Visible = False
    picDownloader.Move cmdDownloader.Left + cmdDownloader.Width + 400, cmdGeneral.Top, picDownloader.Width, picDownloader.Height
    
End Sub

Private Sub cmdGeneral_Click()
    picArrow.Top = cmdGeneral.Top
    picGeneral.Visible = True
    picUpdater.Visible = False
    picDownloader.Visible = False
    picTools.Visible = False
    picLite.Visible = False
    picGeneral.Move cmdGeneral.Left + cmdGeneral.Width + 400, cmdGeneral.Top, picGeneral.Width, picGeneral.Height
End Sub

Private Sub cmdLite_Click()
    picArrow.Top = cmdLite.Top
    picGeneral.Visible = False
    picUpdater.Visible = False
    picDownloader.Visible = False
    picTools.Visible = False
    picLite.Visible = True
    picLite.Move cmdLite.Left + cmdLite.Width + 400, cmdGeneral.Top, picLite.Width, picLite.Height
End Sub

Private Sub cmdRestart_Click()
    Do Until OpenProcess(&H400, 0, Shell(App.Path & "\" & App.EXEName & ".exe")) <> 0
        DoEvents
    Loop
    End
End Sub

Private Sub cmdTools_Click()
    picArrow.Top = cmdTools.Top
    picUpdater.Visible = False
    picDownloader.Visible = False
    picGeneral.Visible = False
    picLite.Visible = False
    picTools.Visible = True
    picTools.Move cmdTools.Left + cmdTools.Width + 400, cmdGeneral.Top, picTools.Width, picTools.Height
    
End Sub

Private Sub cmdUpdater_Click()
    picArrow.Top = cmdUpdater.Top
    picUpdater.Visible = True
    picDownloader.Visible = False
    picGeneral.Visible = False
    picTools.Visible = False
        picLite.Visible = False
    picUpdater.Move cmdUpdater.Left + cmdUpdater.Width + 400, cmdGeneral.Top, picUpdater.Width, picUpdater.Height
End Sub

Private Sub cmdUsage_Click()
    OptiUsage GetCurrentProcess
End Sub

Private Sub Form_Load()
    LoopFormControls Me
    picGeneral.BorderStyle = 0
    picUpdater.BorderStyle = 0
    picDownloader.BorderStyle = 0
    picTools.BorderStyle = 0
    picArrow.BorderStyle = 0
    picLite.BorderStyle = 0
    picArrow.Left = cmdGeneral.Left + cmdGeneral.Width + 100
    LoadResImage 102, "custom", picArrow.hWnd, picArrow.hDC
    picArrow.Refresh
    
    lnDivider.Y2 = Me.Height
    picGeneral.Move cmdGeneral.Left + cmdGeneral.Width + 400, cmdGeneral.Top, picGeneral.Width, picGeneral.Height
    
    swhMinimize.Value = GetIni("General", "MinimizeToTray", True, App.Path & "\YoutubeGrabberOption.ini")
    swhAutoGetLink.Value = GetIni("General", "AutoGetLink", True, App.Path & "\YoutubeGrabberOption.ini")
    swhActivateApp.Value = GetIni("General", "ActivateApp", True, App.Path & "\YoutubeGrabberOption.ini")
    swhAutoUpdate.Value = GetIni("Updater", "AutoUpdate", True, App.Path & "\YoutubeGrabberOption.ini")
    swhAutoOpenLink.Value = GetIni("Updater", "AutoOpenLink", True, App.Path & "\YoutubeGrabberOption.ini")
    swhAutoOpenFolder.Value = GetIni("Downloader", "AutoOpenFolder", False, App.Path & "\YoutubeGrabberOption.ini")
End Sub

Private Sub swhActivateApp_Click(Value As Boolean)
    WriteKey "General", "ActivateApp", Value, App.Path & "\YoutubeGrabberOption.ini"
    If swhActivateApp.Value = True Then swhAutoGetLink.Value = True
End Sub

Private Sub swhAutoGetLink_Click(Value As Boolean)
    WriteKey "General", "AutoGetLink", Value, App.Path & "\YoutubeGrabberOption.ini"
    If swhAutoGetLink.Value = False Then swhActivateApp.Value = False
End Sub

Private Sub swhAutoOpenFolder_Click(Value As Boolean)
    WriteKey "Downloader", "AutoOpenFolder", Value, App.Path & "\YoutubeGrabberOption.ini"
End Sub

Private Sub swhAutoOpenLink_Click(Value As Boolean)
    WriteKey "Updater", "AutoOpenLink", Value, App.Path & "\YoutubeGrabberOption.ini"
End Sub

Private Sub swhAutoUpdate_Click(Value As Boolean)
    WriteKey "Updater", "AutoUpdate", Value, App.Path & "\YoutubeGrabberOption.ini"
End Sub

Private Sub swhMinimize_Click(Value As Boolean)
    WriteKey "General", "MinimizeToTray", Value, App.Path & "\YoutubeGrabberOption.ini"
End Sub


