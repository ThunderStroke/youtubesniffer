Attribute VB_Name = "modURLDecode"
Option Explicit
Private Const E_POINTER As Long = &H80004003

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszURL As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long

Public Function URLDecode( _
    ByVal URL As String, _
    Optional ByVal PlusSpace As Boolean = True) As String
    
    Dim cchUnescaped As Long
    Dim HRESULT As Long
    
    If PlusSpace Then URL = Replace$(URL, "+", " ")
    cchUnescaped = Len(URL)
    URLDecode = String$(cchUnescaped, 0)
    HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    If HRESULT = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    End If
    
    
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

