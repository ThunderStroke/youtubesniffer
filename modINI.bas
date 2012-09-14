Attribute VB_Name = "modINIReader"
'**************************************************************************
'**Ä£ ¿é Ãû£ºINI ÎÄ¼ş²Ù×÷ - ²Ù×÷ini.bas
'**Ëµ    Ã÷£º
'**´´ ½¨ ÈË£ºicecept(Ä§Áé)
'**ÈÕ    ÆÚ£º2008-12-15 00:11:06
'**ĞŞ ¸Ä ÈË£ºicecept(Ä§Áé)
'**ÈÕ    ÆÚ£º2009-1-9 01:20:23
'**Ãè    Êö£ºicecept(Ä§Áé)ÖÆ×÷
'**°æ    ±¾£ºV1.0.0    http://icecept.blog.sohu.com
'*************************************************************************

Option Explicit

'¶Á³ö×Ô¶¨ÒåINIÎÄ¼ş
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Ğ´Èë×Ô¶¨ÒåINIÎÄ¼ş
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'¶Á³ö×Ô¶¨ÒåINIÎÄ¼şÖĞµÄµ¥¸öÇø¶Î¼äµÄËùÓĞ¼üÃûºÍÖµ
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'¶Á³ö×Ô¶¨ÒåINIËùÓĞÇø¶ÎÃû
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'*************************************************************************
'**º¯ Êı Ãû£ºGetIni
'**¹¦ÄÜÃèÊö£º´Ó×Ô¶¨ÒåiniÎÄ¼şÖĞ¶ÁĞÅÏ¢ SectionName:Çø¶ÎÃû KeyWord:¾ßÌå¼üÃû DefString:¶ÁÈ¡Ê§°Üºó·µ»ØÄ¬ÈÏ FileName:iniÎÄ¼şËùÔÚÂ·¾¶
'**²¹    ³ä£º¶Á³öµÄ¶¼ÊÇStringÀàĞÍµÄ,Èç¹ûÄãÏë¶ÁÊı×ÖĞÍµÄ,¿ÉÓÃVBµÄval()º¯Êı½«Æä×ª»¯
'*************************************************************************

Public Function GetIni(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim ResultString As String * 255
    If GetPrivateProfileString(ByVal SectionName, ByVal KeyWord, vbNullString, ByVal ResultString, ByVal Len(ResultString), ByVal FileName) > 0 Then '¹Ø¼ü´ÊµÄÖµ²»Îª¿Õ
        GetIni = Left(ResultString, InStr(ResultString, Chr(0)) - 1)
    Else                                                                        '½«È±Ê¡ÖµĞ´ÈëINIÎÄ¼ş
        WritePrivateProfileString SectionName, KeyWord, DefString, FileName
        GetIni = DefString
    End If
End Function

'*************************************************************************
'**º¯ Êı Ãû£ºGetKeyWord
'**¹¦ÄÜÃèÊö£º´Ó×Ô¶¨ÒåiniÎÄ¼şÖĞµÃµ½¼üÃûºÍÖµ
'**          SectionName:Çø¶ÎÃû  DefString:Ä¬ÈÏÖµ  FileName:iniÎÄ¼şËùÔÚÂ·¾¶
'**²¹    ³ä£ºÿ
'*************************************************************************
Public Function GetKeyWord(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255
    If GetPrivateProfileSection(ByVal SectionName, ByVal szBuf, Len(szBuf), ByVal FileName) > 0 Then
        'Í¬Ê±»ñÈ¡¼üÃûºÍÖµ
        GetKeyWord = Left(szBuf, InStr(szBuf, Chr(0)) - 1)
    Else
        WritePrivateProfileString SectionName, DefString, vbNullString, FileName
        GetKeyWord = DefString
    End If
End Function

'*************************************************************************
'**º¯ Êı Ãû£ºGetKey
'**¹¦ÄÜÃèÊö£º´Ó×Ô¶¨ÒåiniÎÄ¼şÖĞµÃµ½¼üÃû
'**          SectionName:Çø¶ÎÃû  DefString:Ä¬ÈÏÖµ  FileName:iniÎÄ¼şËùÔÚÂ·¾¶
'**²¹    ³ä£ºÿ
'*************************************************************************
Public Function GetKey(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255, ResultString As String
    ResultString = GetKeyWord(ByVal SectionName, ByVal szBuf, ByVal FileName)
    If InStr(ResultString, "=") <> 0 Then
        GetKey = Left(ResultString, InStr(ResultString, "=") - 1)
    Else
        GetKey = DefString
    End If
End Function

Public Function WriteKey(ByVal Çø¶ÎÃû As String, ByVal ¼üÃû As String, ByVal Öµ As String, ByVal ÎÄ¼şÃû As String) As String
    WritePrivateProfileString Çø¶ÎÃû, ¼üÃû, Öµ, ÎÄ¼şÃû
End Function

Public Function DelKey(ByVal Çø¶ÎÃû As String, ByVal ¼üÃû As String, ByVal ÎÄ¼şÃû As String) As String
    WritePrivateProfileString "Çø¶ÎÃû", "¼üÃû", vbNullString, ÎÄ¼şÃû
End Function
    'WritePrivateProfileString "a", "b", vbNullString, App.Path & "\test.ini"

'Éú³ÉµÄtest.ini ÎÄ¼ş¸ñÊ½:
'
'[a]              a: Çø¶ÎÃû
'b=c¡¡    ¡¡¡¡¡¡¡¡b£º¼üÃû  c£ºÖµ
'[d]              d: Çø¶ÎÃû
'e=f              e£º¼üÃû  f£ºÖµ
'--------------------------------------------------------------------------------------
'²Ù×÷iniÎÄ¼şµÄº¯ÊıÁĞ±í:
'Ö»ÄÜ²Ù×÷win.iniµÄº¯Êı
'WriteProfileString ÔÚwin.ini³õÊ¼»¯ÎÄ¼şÖ¸¶¨Ğ¡½ÚÄÚÉèÖÃÒ»¸ö×Ö´®
'WriteProfileSection Îªwin.ini³õÊ¼»¯ÎÄ¼şÖĞÒ»¸öÖ¸¶¨µÄĞ¡½ÚÉèÖÃËùÓĞÏîÃûºÍÖµ
'GetProfileInt È¡µÃwin.ini³õÊ¼»¯ÎÄ¼şÖĞÖ¸¶¨ÌõÄ¿µÄÒ»¸öÕûÊıÖµ
'GetProfileSection               »ñÈ¡Ö¸¶¨Ğ¡½Ú£¨ÔÚwin.iniÎÄ¼şÖĞ£©ËùÓĞÏîÃûºÍÖµµÄÒ»¸öÁĞ±í
'GetProfileString Îªwin.ini³õÊ¼»¯ÎÄ¼şÖĞÖ¸¶¨µÄÌõÄ¿È¡µÃ×Ö´®
'======================================================================================
'¼´ÄÜ²Ù×÷win.iniµÄº¯Êı£¬ÓÖÄÜ²Ù×÷×Ô¶¨ÒåiniÎÄ¼şµÄº¯Êı
'WritePrivateProfileString ÔÚ³õÊ¼»¯ÎÄ¼şÖ¸¶¨Ğ¡½ÚÄÚÉèÖÃÒ»¸ö×Ö´®
'WritePrivateProfileSection      ÎªÒ»¸ö³õÊ¼»¯ÎÄ¼ş£¨.ini£©ÖĞÖ¸¶¨µÄĞ¡½ÚÉèÖÃËùÓĞÏîÃûºÍÖµ
'GetPrivateProfileInt Îª³õÊ¼»¯ÎÄ¼şÖĞÖ¸¶¨µÄÌõÄ¿»ñÈ¡Ò»¸öÕûÊıÖµ
'GetPrivateProfileSection »ñÈ¡Ö¸¶¨Ğ¡½ÚËùÓĞÏîÃûºÍÖµµÄÒ»¸öÁĞ±í
'GetPrivateProfileString Îª³õÊ¼»¯ÎÄ¼şÖĞÖ¸¶¨µÄÌõÄ¿È¡µÃ×Ö´®
'GetPrivateProfileSectionNames ¶Á³öINIÎÄ¼şÖĞËùÓĞÇø¶ÎÃû

'Write
'      Çø¶ÎÃû£¬¼üÃû£¬Öµ£¬ÎÄ¼şÃû
    'WritePrivateProfileString "a", "b", "c", App.Path & "\test.ini"
    'WritePrivateProfileString "d", "e", "f", App.Path & "\test.ini"
 'ĞŞ¸Ä¼üÖµ
    'Dim AA As String
    'AA = InputBox("ÇëÊäÈëĞÂµÄaÇø¶Î¼üÃûbÏÂµÄÖµ", Me.Caption)
    'If AA <> vbNullString Then
        'WritePrivateProfileString "a", "b", AA, App.Path & "\test.ini"
        'MsgBox "ĞŞ¸Ä³É¹¦"
    'End If
    
    'É¾³ıaÇø¶ÎÏÂµÄÒ»¸ö¼üÃû,É¾³ı¼üÃûºó£¬¼üÃûÏÂµÄÖµÒ²»áÍ¬Ê±É¾³ı
    '                        Çø¶ÎÃû ¼üÃû
    'WritePrivateProfileString "a", "b", vbNullString, App.Path & "\test.ini"
'É¾³ıÕû¸öÇø¶Î,,É¾³ıÇø¶Îºó£¬Çø¶ÎÏÂµÄ¼üÃûºÍÖµÒ²»áÍ¬Ê±É¾³ı
    '        Çø¶ÎÃû
    'WritePrivateProfileString "a", vbNullString, vbNullString, App.Path & "\test.ini"
    'WritePrivateProfileString "d", vbNullString, vbNullString, App.Path & "\test.ini"

'¶ÁÈ¡Êı¾İ       Çø¶ÎÃû     ¼üÃû      Ä¬ÈÏÖµ             ÎÄ¼şÃû
    'MsgBox GetIni("rBnwares", "Year", vbNullString, "c:\windows\win.ini")

'»ñÈ¡¼üÃû
    'Dim szBuf As String * 255
    '                Çø¶ÎÃû£¬     ¼üÃû£¬    ÎÄ¼şÃû
    'MsgBox GetKeyWord("rBnwares", szBuf, "c:\windows\win.ini")
 '»ñÈ¡¼üÃû
    'Dim szBuf As String * 255
    '                Çø¶ÎÃû£¬   ¼üÃû£¬    ÎÄ¼şÃû
    'MsgBox GetKey("rBnwares", szBuf, "c:\windows\win.ini")

