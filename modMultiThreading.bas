Attribute VB_Name = "modMultiThreading"
'5/5/2010 base on vb multithread dll(code by izero@Slovakia)
'vb multithread without tlb,pcode,dll,ax,and so on (v1.1) by download@vbgood
'Modify by GaryNg@vbgood@4/9/2012
'From http://www.vbgood.com/thread-93124-1-1.html
Private Type mIID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long
Private modnFakeHeader As Long
Private Declare Function VBGetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModName As Long) As Long
Private Declare Sub UserDllMain Lib "msvbvm60.dll" (u1 As Long, u2 As Long, ByVal u3_h As Long, ByVal u4_1 As Long, ByVal u5_0 As Long)
Private Declare Function VBDllGetClassObject Lib "msvbvm60.dll" (g1 As Long, g2 As Long, ByVal g3_vbHeader As Long, REFCLSID As Long, REFIID As mIID, ppv As Long) As Long
Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
Private Declare Sub CoUninitialize Lib "ole32.dll" ()

Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long


Sub InitFakeHeader()
    Dim nFreeFile As Long
    Dim strMyFileName As String
    Dim bytFakeHeader() As Byte
    Dim bytVBID() As Byte
    strMyFileName = App.Path & "\" & App.EXEName & ".exe"
    nFreeFile = FileLen(strMyFileName)
    If nFreeFile > 0 Then
        ReDim bytFakeHeader(nFreeFile - 1)
        nFreeFile = FreeFile
        Open strMyFileName For Binary As nFreeFile
        Get nFreeFile, , bytFakeHeader
        Close nFreeFile
        bytVBID = StrConv("VB5!", vbFromUnicode)
        modnFakeHeader = InStrB(bytFakeHeader, bytVBID)
    End If
    'vb multithread without tlb,pcode,dll,ax,and so on (v1.1) by download@vbgood
    'this sub should be call only once at app start.
End Sub

Sub InitVBdll()
    Dim pIID As mIID
    Dim pDummy As Long
    pIID.data1 = 1
    pIID.data4(0) = &HC0
    pIID.data4(7) = &H46
    Dim u1 As Long, u2 As Long, u3 As Long
    Dim g3 As Long
    u3 = VBGetModuleHandle(0)
    Dim pVBHeader As Long
    UserDllMain u1, u2, u3, 1, 0
    If modnFakeHeader > 0 Then
        g3 = u3 + modnFakeHeader - 1
        VBDllGetClassObject u1, u2, g3, pDummy, pIID, pDummy
    End If
End Sub

Public Sub Thread1()
    Dim hr As Long
    CreateIExprSrvObj 0, 4, 0
    hr = CoInitialize(0)
    Call InitVBdll
    Call Form2.XXX
    Call CoUninitialize
    
End Sub


Public Sub Thread2()
    Dim hr As Long
    CreateIExprSrvObj 0, 4, 0
    hr = CoInitialize(0)
    Call InitVBdll
    Call Form2.XXXX
    Call CoUninitialize
    
End Sub

