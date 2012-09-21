VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   5820
   ClientLeft      =   2295
   ClientTop       =   1590
   ClientWidth     =   5130
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   4440
         Width           =   4725
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
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
      Height          =   4800
      Left            =   120
      Picture         =   "frmAbout.frx":5F32
      ScaleHeight     =   4800
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label Label2 
      Caption         =   "Blog: http://garyngzhongbo.blopgspot.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmAbout.frx":69F74
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Google Page: http://code.google.com/p/youtubesniffer/"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmAbout.frx":6A0C6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5040
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MHeight As Long
Dim MWidth As Long

Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal Stream As Any, Image As Long) As GpStatus
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)


Private Sub Form_Activate()
    'Do While Me.Height < MHeight
    'Me.Height = Me.Height + MHeight / 110
    'Me.Width = Me.Width + MWidth / 110
    'SetWindowPos Me.hwnd, 0, (Screen.Width - Me.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - Me.Height) / 2 / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
    'DoEvents
    'Loop
    Dim nTime As Long
    nTime = Timer()
    Do
        DoEvents
        If Timer() - nTime > 5 Then Unload Me
    Loop
    
End Sub

Private Sub Form_Load()
    'MHeight = Me.Height
    'MWidth = Me.Width
    RoundRect Me.hDC, 0, 0, Me.Width / Screen.TwipsPerPixelX - 1, Me.Height / Screen.TwipsPerPixelY - 1, 20, 20
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 20, 20), True
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Current Version: Build " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    'lblDescription.Caption = App.Title & " is a software that build for download Youtube's video. It has many function that other downloader did not exist such as getting the video information, it also have many option for selecting video download link and other useful function. Brought to you by GaryNg - A member of AG Dev Team"
    
    'Me.Height = 1
    'Me.Width = 1
    
    Call LoadResImage(101, "custom", Picture1.hWnd, Picture1.hDC)
    Picture1.Refresh
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Do While Me.Height > 1
    'Me.Height = Me.Height - MHeight / 110
    'Me.Width = Me.Width - MWidth / 110
    'SetWindowPos Me.hwnd, 0, (Screen.Width - Me.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - Me.Height) / 2 / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
    'DoEvents
    'Loop
End Sub

Private Sub Label1_Click()
    VisitURL "http://code.google.com/p/youtubesniffer/"
End Sub

Private Sub Label2_Click()
    VisitURL "http://garyngzhongbo.blopgspot.com"
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub


