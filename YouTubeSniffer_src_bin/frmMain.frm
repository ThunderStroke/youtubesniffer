VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "YouTube Sniffer Tools"
   ClientHeight    =   7815
   ClientLeft      =   2145
   ClientTop       =   1965
   ClientWidth     =   15630
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7815
   ScaleWidth      =   15630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "split download link from webpage"
      Height          =   7335
      Left            =   11400
      TabIndex        =   11
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdWebSplit 
         Caption         =   "Split Video Download Link from Webpage"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtWebDL 
         Height          =   6255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "get_video_info split download link"
      Height          =   7335
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtDonwloadLink 
         Height          =   6255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton cmdSplit 
         Caption         =   "Split Video Download Link from get_video_info"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "get_video_info"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   6975
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtRet 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "YouTube Video Link:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Press Enter When You're Done"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "URL Decoder"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdDecode 
         Caption         =   "Decode"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   6735
      End
      Begin VB.TextBox txtCode 
         Height          =   4215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDecode_Click()
    txtCode.Text = UTF8_UrlDecode(UTF8_UrlDecode(txtCode.Text))
End Sub

Private Sub cmdSplit_Click()
    If txtCode.Text = "" Then
        MsgBox "Please decode the get_video_file in the URL decoder first!"
        txtCode.SetFocus
    Else
        txtDonwloadLink.Text = ""
        Dim strDownloadLink() As String
        Dim nI As Long
        strDownloadLink = Split(txtCode.Text, "url=")
        For nI = 0 To UBound(strDownloadLink)
            If InStr(strDownloadLink(nI), "o-o---preferred---") > 0 Then
                DoEvents
                If InStr(strDownloadLink(nI), ";") > 0 Then
                    txtDonwloadLink = txtDonwloadLink & Mid(strDownloadLink(nI), 1, InStr(strDownloadLink(nI), ";") - 1) & vbCrLf & "=========================" & vbCrLf & vbCrLf
                Else
                    txtDonwloadLink = txtDonwloadLink & Mid(strDownloadLink(nI), 1, Len(strDownloadLink(nI)) - 1 - Len(strDownloadLink(nI)) + InStrRev(strDownloadLink(nI), "&itag=")) & vbCrLf & "=========================" & vbCrLf & vbCrLf
                End If
            End If
        Next
        
    End If
End Sub

Private Sub cmdWebSplit_Click()
    If txtCode.Text = "" Then
        MsgBox "Did you copy the ""var swf = "" and decode it from the webpage?"
        txtCode.SetFocus
    Else
        txtWebDL.Text = ""
        Dim strDownloadLink() As String
        'Dim strWebCode As String
        Dim nI As Long
        'strWebCode = Mid(txtCode.Text, InStr(txtCode.Text, "url_encoded_fmt_stream_map"), Len(txtCode.Text))
        strDownloadLink = Split(Mid(txtCode.Text, InStr(txtCode.Text, "url_encoded_fmt_stream_map"), Len(txtCode.Text)), "url=")
        For nI = 0 To UBound(strDownloadLink)
            If InStr(strDownloadLink(nI), "o-o---preferred---") > 0 Then
                DoEvents
                If InStr(strDownloadLink(nI), ";") > 0 Then
                    txtWebDL.Text = txtWebDL.Text & Mid(strDownloadLink(nI), 1, InStr(strDownloadLink(nI), ";") - 1) & Replace(Mid(Mid(strDownloadLink(nI), InStr(Replace(strDownloadLink(nI), "&sig=", "&signature="), "&signature=") + 1, Len(strDownloadLink(nI))), 1, InStr(Mid(strDownloadLink(nI), InStr(Replace(strDownloadLink(nI), "&sig=", "&signature="), "&signature=") + 1, Len(strDownloadLink(nI))), "&") - 1), "sig=", "&signature=") & vbCrLf & vbCrLf & "====================" & vbCrLf
                Else
                    txtWebDL.Text = txtWebDL.Text & Replace(Mid(strDownloadLink(nI), 1, InStrRev((strDownloadLink(nI)), ",") - 1), "&sig=", "&signature=") & vbCrLf & vbCrLf & "====================" & vbCrLf
                End If
                'txtWebDL.Text = txtWebDL.Text & InStr(Replace(strDownloadLink(nI), "&sig=", "&signature="), "&signature=") & vbCrLf & vbCrLf & "^^^^====================^^^" & vbCrLf
                
            End If
        Next
        
        'Debug.Print strWebCode
    End If
    
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strWebHtml As String
    If KeyCode = 13 Then
        txtRet.Text = ""
        txtRet.Text = "[Video_id:] " & ExtractMatch(txtLink.Text, "v=([A-Za-z0-9-_]+)") & vbCrLf
        txtRet.Text = txtRet.Text & "[get_video_info:] http://youtube.com/get_video_info?video_id=" & ExtractMatch(txtLink.Text, "v=([A-Za-z0-9-_]+)") & vbCrLf
    End If
    
End Sub

Private Function ExtractMatch(Text, Pattern)
    Dim Regex As Object
    Dim Matches
    
    Set Regex = New RegExp
    Regex.Pattern = Pattern
    
    Set Matches = Regex.Execute(Text)
    If Matches.Count = 0 Then
        ExtractMatch = ""
        Exit Function
    End If
    
    ExtractMatch = Matches(0).SubMatches(0)
End Function

Private Function UTF8_UrlDecode(ByVal URL As String)
    Dim B, ub
    Dim UtfB
    Dim UtfB1, UtfB2, UtfB3
    Dim i, n, s
    n = 0
    ub = 0
    For i = 1 To Len(URL)
        B = Mid(URL, i, 1)
        Select Case B
        Case "+"
            s = s & " "
        Case "%"
            ub = Mid(URL, i + 1, 2)
            UtfB = Val("&H" & ub)
            If UtfB < 128 Then
                i = i + 2
                s = s & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000
                UtfB2 = (CInt("&H" & Mid(URL, i + 4, 2)) And &H3F) * &H40
                UtfB3 = CInt("&H" & Mid(URL, i + 7, 2)) And &H3F
                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                i = i + 8
            End If
        Case Else
            s = s & B
        End Select
    Next
    UTF8_UrlDecode = s
    
End Function
