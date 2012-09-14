VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   5475
   ClientLeft      =   2295
   ClientTop       =   1590
   ClientWidth     =   8610
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
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   8415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   225
         Left            =   0
         TabIndex        =   4
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   6840
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4770
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Dim MHeight As Long
Dim MWidth As Long

Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal Stream As Any, Image As Long) As GpStatus
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'Do While Me.Height < MHeight
    'Me.Height = Me.Height + MHeight / 110
    'Me.Width = Me.Width + MWidth / 110
    'SetWindowPos Me.hwnd, 0, (Screen.Width - Me.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - Me.Height) / 2 / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
    'DoEvents
    'Loop
    
End Sub



Private Sub Form_Load()
    'MHeight = Me.Height
    'MWidth = Me.Width
    
    
    RoundRect Me.hDC, 0, 0, Me.Width / Screen.TwipsPerPixelX - 1, Me.Height / Screen.TwipsPerPixelY - 1, 20, 20
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Current Version: Build " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.Title & " is a software that build for download Youtube's video. It has many function that other downloader did not exist such as getting the video information, it also have many option for selecting video download link and other useful function. Brought to you by GaryNg - A member of AG Dev Team"
    
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 20, 20), True
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

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hWnd
End Sub

