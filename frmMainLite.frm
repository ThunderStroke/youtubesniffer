VERSION 5.00
Begin VB.Form frmMainLite 
   Caption         =   "AG Youtube Video Grabber"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGo 
      Caption         =   "OK"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtLink 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Link:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMainLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

