VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YouTube Video Sniffer"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   16560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   6735
      Left            =   12000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox txtInfo 
      Height          =   6735
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.ListBox lstDownloadLink 
      Height          =   2085
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   120
      Top             =   4560
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UNHOOK"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser wbHook 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      ExtentX         =   9975
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strGen204 As String

Private Sub Command4_Click()
    txtLink.Text = "http://www.youtube.com/watch?v=-Djj6pfR9KU&feature=relmfu"
    txtLink_KeyDown 13, 1
End Sub

Private Sub Form_Load()
    wbHook.Silent = True
    wbHook.Navigate "about:blank"
    OptiRAM
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub tmrCheck_Timer()
    Dim i As Long
    If SafeArrayGetDim(strSendInfo) = 0 Then Exit Sub
    For i = 0 To UBound(strSendInfo)
        DoEvents
        If InStr(strSendInfo(i), "generate_204") > 0 Then
            strGen204 = strSendInfo(i)
            wbHook.Stop
            UnHook
            tmrCheck.Enabled = False
            Exit Sub
        End If
    Next
End Sub


Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Hook                                                                    'It will unhook when recieved GET info in tmrCheck_Timer
        wbHook.Navigate txtLink.Text
        Do Until strGen204 <> ""
            DoEvents
        Loop
        strGen204 = URLDecode(strGen204)
        LoadInfo wbHook, txtInfo
    End If
End Sub

