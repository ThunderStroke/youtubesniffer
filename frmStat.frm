VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program Status"
   ClientHeight    =   1065
   ClientLeft      =   1725
   ClientTop       =   10065
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstState 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pWndProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WindowProc_Child
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuShowStat.Checked = False
    SetWindowLong Me.hwnd, GWL_WNDPROC, pWndProc
    
End Sub


