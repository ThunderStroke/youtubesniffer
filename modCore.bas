Attribute VB_Name = "modCore"
Option Explicit
'Core function module

Public Function SeperateSWF(ByVal strVideoURL As String) As String
    Dim strWebHTML As String
    Dim strVideoLink As String
    strVideoLink = strVideoURL
    If LCase(Left(strVideoURL, 7)) <> "http://" Then strVideoLink = "http://" & strVideoLink
    strWebHTML = DownHTML(strVideoLink, "UTF-8")
    Debug.Print strWebHTML
End Function

