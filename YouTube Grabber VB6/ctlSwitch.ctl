VERSION 5.00
Begin VB.UserControl cltSwitch 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   MaskColor       =   &H000000FF&
   MaskPicture     =   "ctlSwitch.ctx":0000
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1800
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Index           =   0
      Left            =   1440
      Picture         =   "ctlSwitch.ctx":A9E2
      ScaleHeight     =   2760
      ScaleWidth      =   1170
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Image MaskPic 
      Height          =   345
      Index           =   1
      Left            =   120
      Picture         =   "ctlSwitch.ctx":153C4
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Image MaskPic 
      Height          =   345
      Index           =   0
      Left            =   240
      Picture         =   "ctlSwitch.ctx":15B48
      Top             =   2520
      Width           =   1170
   End
End
Attribute VB_Name = "cltSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SwitchEnable As Boolean '有效无效
Dim SwitchCondition As Integer '0普通1指向2按下
Dim SwitchValue As Boolean '启用关闭
Dim newValue As Boolean

Dim SwitchAim As Boolean '鼠标在控件上
Enum mySkin
    Safe = 0
    Kingsoft = 1
    KuGou = 2
End Enum
Dim SwitchSkin As mySkin

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Type POINT_API
    X As Long
    Y As Long
End Type

Public Event Click(Value As Boolean)
Public Event MouseOut()
Public Event MouseIn()

Private Sub UserControl_Initialize()
    Call Refresh
End Sub
Private Sub UserControl_InitProperties()
    SwitchCondition = 0
    SwitchEnable = True
    SwitchValue = True
    SwitchAim = False
    SwitchSkin = Safe
End Sub
Private Sub Refresh() '0普通1指向2按下
    If SwitchSkin = Safe Then
        UserControl.PaintPicture Picture1(0).Picture, 0, 0, 78, 23, 0, SwitchCondition * 23, 78, 23
        Set UserControl.MaskPicture = MaskPic(IIf(SwitchValue, 0, 1)).Picture
    End If
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And SwitchEnable Then
        newValue = Not SwitchValue
        SwitchCondition = IIf(SwitchValue, 2, 6)
        Call Refresh
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And SwitchEnable And SwitchAim Then
        SwitchValue = newValue
        SwitchCondition = IIf(SwitchValue, 1, 5)
        Call Refresh
        RaiseEvent Click(SwitchValue)
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Timer1.Enabled Then
    End If
End Sub
Private Sub Timer1_Timer()
    If Not SwitchEnable Then Exit Sub
    Dim dot As POINT_API
    Call GetCursorPos(dot)
    ScreenToClient UserControl.hWnd, dot
    If dot.X < UserControl.ScaleLeft Or dot.Y < UserControl.ScaleTop Or _
            dot.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        If SwitchAim Then '移出
            SwitchAim = False
            SwitchCondition = IIf(SwitchValue, 0, 4)
            Call Refresh
            RaiseEvent MouseOut
        End If
    Else
        If Not SwitchAim Then
            SwitchAim = True
            SwitchCondition = IIf(SwitchValue, 1, 5)
            Call Refresh
            RaiseEvent MouseIn
        End If
    End If
End Sub
Public Property Get Enable() As Boolean
    Enable = SwitchEnable
End Property
Public Property Let Enable(ByVal vNewValue As Boolean)
    SwitchEnable = vNewValue
    
    If SwitchEnable Then
        If SwitchAim Then
            SwitchCondition = IIf(SwitchValue, 1, 5)
        Else
            SwitchCondition = IIf(SwitchValue, 0, 4)
        End If
    Else
        SwitchCondition = IIf(SwitchValue, 3, 7)
    End If
    Call Refresh
End Property
Public Property Get Value() As Boolean
    Value = SwitchValue
End Property
Public Property Let Value(ByVal vNewValue As Boolean)
    SwitchValue = vNewValue
    
    If SwitchEnable Then
        If SwitchAim Then
            SwitchCondition = IIf(SwitchValue, 1, 5)
        Else
            SwitchCondition = IIf(SwitchValue, 0, 4)
        End If
    Else
        SwitchCondition = IIf(SwitchValue, 3, 7)
    End If
    Call Refresh
    RaiseEvent Click(SwitchValue)
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("value", True)
    Enable = PropBag.ReadProperty("enable", True)
    Timer1.Enabled = Ambient.UserMode
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("value", SwitchValue, True)
    Call PropBag.WriteProperty("enable", SwitchEnable, True)
End Sub

Private Sub UserControl_Resize()
    If SwitchSkin = Safe Then
        UserControl.Width = 78 * 15
        UserControl.Height = 23 * 15
    End If
End Sub


