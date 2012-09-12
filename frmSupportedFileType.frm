VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupportedFileType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supported File Type"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupportedFileType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Search By Dragging File In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   5175
      Begin MSComctlLib.ListView lvwSFileFormat 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Directory"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Is Supported"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Format"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search By Entering File Format (*.*)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   4575
      Begin VB.CommandButton cmdCheckFileByType 
         Caption         =   "Check Now!"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtSFileFormat 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Format(Extension):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search For Supported File Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   10935
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "OR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supported Video/Audio Format"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   10935
      Begin VB.TextBox txtFormatList 
         Height          =   2535
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frmSupportedFileType.frx":5F32
         Top             =   360
         Width           =   10575
      End
   End
End
Attribute VB_Name = "frmSupportedFileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FsoSys As New FileSystemObject


Private Sub cmdCheckFileByType_Click()
    txtSFileFormat.Text = Replace(txtSFileFormat.Text, ".", "")
    txtSFileFormat.Text = Replace(txtSFileFormat.Text, "*", "")
    
    If InStr(txtFormatList.Text, "*." & txtSFileFormat.Text & " ") > 0 Then
        lblStat.Caption = "Found"
        txtFormatList.SetFocus
        txtFormatList.SelStart = InStr(txtFormatList.Text, "*." & txtSFileFormat.Text & " ") - Len("*.")
        Dim tmp As String
        tmp = Right(txtFormatList.Text, Len(txtFormatList.Text) - InStr(txtFormatList.Text, "*." & txtSFileFormat.Text & " ") + Len("*."))
        txtFormatList.SelLength = InStr(tmp, vbCrLf)
    Else
        lblStat.Caption = "Not Found"
    End If
End Sub

Private Sub lvwSFileFormat_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i&
    Dim lx As ListItem
    
    If Data.GetFormat(vbCFFiles) Then
        ReDim s(1 To Data.Files.Count)
        For i = 1 To Data.Files.Count
            If PathIsDirectory(Data.Files(i)) <> FILE_ATTRIBUTE_DIRECTORY Then
                DoEvents
                
                Set lx = lvwSFileFormat.ListItems.Add(, , Data.Files(i))
                
                If InStr(txtFormatList.Text, "*." & FsoSys.GetExtensionName(Data.Files(i)) & " ") Then
                    
                    lx.SubItems(2) = FsoSys.GetExtensionName(Data.Files(i))
                    
                    lx.SubItems(1) = "Yes"
                Else
                    lx.SubItems(1) = "No"
                End If
            End If
        Next
    End If
End Sub


Private Sub txtSFileFormat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCheckFileByType_Click
End Sub
