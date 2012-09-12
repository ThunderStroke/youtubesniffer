Attribute VB_Name = "modFunc"
Option Explicit
'Array Uninitialized
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'Array Uninitialized===
'Optimize RAM Usage
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Optimize RAM Usage===

Public strDecodedSWF As String

Public Sub LoadInfo(ByVal wbSource As WebBrowser, ByVal txtDest As TextBox)
    Dim nI As Long
    On Error Resume Next
    If Len(txtDest.Text) > 0 Then txtDest.Text = ""
    For nI = 1 To wbSource.Document.All.Length - 1
        If LCase(wbSource.Document.All(nI).tagName) = "title" Then
            txtDest.Text = txtDest.Text & "[Title:] " & wbSource.Document.All(nI).innerText & vbCrLf
        End If
        
        If LCase(wbSource.Document.All(nI).tagName) = "meta" And LCase(wbSource.Document.All(nI).Name) = "attribution" Then
            txtDest.Text = txtDest.Text & "[Author:] " & Mid(wbSource.Document.All(nI).Content, 1, InStr(wbSource.Document.All(nI).Content, "%2Buser") - 1) & vbCrLf
            txtDest.Text = txtDest.Text & "[Author Channel:] " & "http://www.youtube.com/user/" & Mid(wbSource.Document.All(nI).Content, 1, InStr(wbSource.Document.All(nI).Content, "%2Buser") - 1) & vbCrLf
        End If
        
        If LCase(wbSource.Document.All(nI).tagName) = "span" And LCase(wbSource.Document.All(nI).ClassName) = "watch-view-count" Then
            txtDest.Text = txtDest.Text & "[Views:] " & wbSource.Document.All(nI).innerText & vbCrLf
        End If
        
        If LCase(wbSource.Document.All(nI).tagName) = "span" And LCase(wbSource.Document.All(nI).ClassName) = "likes" Then
            txtDest.Text = txtDest.Text & "[Likes:] " & wbSource.Document.All(nI).innerText & vbCrLf
        End If
        
        If LCase(wbSource.Document.All(nI).tagName) = "span" And LCase(wbSource.Document.All(nI).ClassName) = "dislikes" Then
            txtDest.Text = txtDest.Text & "[Dislikes:] " & wbSource.Document.All(nI).innerText & vbCrLf
        End If
        If LCase(wbSource.Document.All(nI).tagName) = "script" And InStr(LCase(wbSource.Document.All(nI).innerHTML), "var swf = ") Then
            strDecodedSWF = URLDecode(URLDecode(URLDecode(wbSource.Document.All(nI).innerHTML)))
        End If
        
        DoEvents
    Next
    'Text2.Text = Text2.Text & "<" & i & ">##########################################" & vbCrLf
    'Text2.Text = Text2.Text & "Name: " & wbSource.Document.All(ni).Name & vbCrLf
    'Text2.Text = Text2.Text & "TagName: " & wbSource.Document.All(ni).tagName & vbCrLf
    'Text2.Text = Text2.Text & "Type: " & wbSource.Document.All(ni).Type & vbCrLf
    'Text2.Text = Text2.Text & "Src: " & wbSource.Document.All(ni).src & vbCrLf
    'Text2.Text = Text2.Text & "Value: " & wbSource.Document.All(ni).Value & vbCrLf
    'Text2.Text = Text2.Text & "Href:" & wbSource.Document.All(ni).href & vbCrLf
    'Text2.Text = Text2.Text & "ID: " & wbSource.Document.All(ni).Id & vbCrLf
    'Text2.Text = Text2.Text & "OuterText: " & wbSource.Document.All(ni).outerText & vbCrLf
    'Text2.Text = Text2.Text & "Class: " & wbSource.Document.All(ni).ClassName & vbCrLf
    
End Sub


Public Sub OptiRAM()
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1
End Sub

