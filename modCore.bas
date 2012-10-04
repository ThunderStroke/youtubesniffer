Attribute VB_Name = "modCore"
Option Explicit
'Core function module
Public strWebHTML As String
Public strDecodedSWF As String

Public Function SeperateSWF(ByVal strVideoURL As String) As String
    Dim strVideoLink As String
    strVideoLink = strVideoURL
    If LCase(Left(strVideoURL, 7)) <> "http://" Then strVideoLink = "http://" & strVideoLink
    
    strWebHTML = DownHTML(strVideoLink, "UTF-8")
    strWebHTML = Replace(strWebHTML, Chr(9), "")
    strWebHTML = Replace(strWebHTML, vbCrLf, "")
    strWebHTML = Replace(strWebHTML, vbLf, "")
    strWebHTML = Replace(strWebHTML, vbCr, "")
    If strWebHTML = "" Then Exit Function
    strDecodedSWF = URLDecode(URLDecode(URLDecode(Mid(Mid(strWebHTML, InStr(strWebHTML, "var swf ="), InStr(strWebHTML, ".innerHTML = swf;") - InStr(strWebHTML, "var swf =")), InStr(Mid(strWebHTML, InStr(strWebHTML, "var swf ="), InStr(strWebHTML, ".innerHTML = swf;") - InStr(strWebHTML, "var swf =")), "url_encoded_fmt_stream_map")))))
    'frmMain.Text1.Text = strDecodedSWF
End Function

Public Function LoadVideoInfo(ByVal strVideoLink As String, ByVal txtVideoID As TextBox, ByVal txtVideoTitle As TextBox, ByVal txtVideoViews As TextBox, ByVal txtVideoLength As TextBox, ByVal txtVideoUploader As TextBox, ByVal txtVideoChannel As TextBox, ByVal picVideoScreenShotCon As PictureBox, ByVal picVideoScreenShotView As PictureBox)
    txtVideoID.Text = LoadVideoID(strVideoLink)
    txtVideoTitle.Text = LoadVideoTitle
    txtVideoViews.Text = LoadVideoViews
    txtVideoLength.Text = LoadVideoLength
    Dim strVideoUploader As String, strVideoChannel As String
    LoadAuthorInfo strVideoUploader, strVideoChannel
    txtVideoUploader.Text = strVideoUploader
    txtVideoChannel.Text = strVideoChannel
    LoadPicScreenShot picVideoScreenShotCon, picVideoScreenShotView
End Function

Public Function LoadVideoInfoLite(ByVal strVideoLink As String, ByRef strVideoTitle As String, ByRef strVideoUploader As String, ByRef strVideoChannel As String, ByRef strVideoID As String, ByRef strVideoViews As String, ByRef strVideoLength As String)
    strVideoTitle = LoadVideoTitle
    LoadAuthorInfo strVideoUploader, strVideoChannel
    strVideoID = LoadVideoID(strVideoLink)
    strVideoViews = LoadVideoViews
    strVideoLength = LoadVideoLength
End Function

Public Function LoadVideoID(ByVal strVideoLink As String)
    LoadVideoID = ExtractMatch(strVideoLink, "v=([A-Za-z0-9-_]+)")
End Function

Public Function LoadVideoTitle()
    LoadVideoTitle = Mid(strWebHTML, InStr(strWebHTML, "<title>") + Len("<title>"), InStr(strWebHTML, "</title>") - InStr(strWebHTML, "<title>") - 7)
End Function

Public Function LoadVideoViews()
    LoadVideoViews = Mid(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), InStr(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), "<strong>") + Len("<strong>"), InStr(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), "</strong>") - Len("</strong>") - 6)
End Function

Public Function LoadVideoLength()
    LoadVideoLength = ConvertToTime(Val(Mid(strWebHTML, InStr(strWebHTML, ";length_seconds=") + Len(";length_seconds="))))
End Function

Public Function LoadAuthorInfo(ByRef strVideoUploader As String, ByRef strVideoChannel As String)
    Dim strTmp As String
    strTmp = Trim(Mid(strWebHTML, InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype="""), InStr(InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype="""), strWebHTML, "</span>") - InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype=""")))
    strVideoChannel = Left(Mid(strTmp, InStr(strTmp, "<link itemprop=""url"" href=""") + Len("<link itemprop=""url"" href=""")), Len(Mid(strTmp, InStr(strTmp, "<link itemprop=""url"" href=""") + Len("<link itemprop=""url"" href="""))) - 2)
    strVideoUploader = Replace(strVideoChannel, "http://www.youtube.com/user/", "")
End Function

Public Function LoadPicScreenShot(ByVal picVideoScreenShot As PictureBox, ByVal picVideoScreenShotView As PictureBox)
    Set picVideoScreenShot.Picture = LoadPicture(Mid(Mid(strWebHTML, InStr(strWebHTML, "<meta property=""og:image"" content=") + Len("<meta property=""og:image"" content=")), 2, InStr(Mid(strWebHTML, InStr(strWebHTML, "<meta property=""og:image"" content=") + Len("<meta property=""og:image"" content=")), """>") - 2))
    FitPictureToBox picVideoScreenShot, picVideoScreenShotView
End Function

Public Function ProcessDownloadLinks(ByRef strDownloadLink() As String)
    Dim nI As Long
    Dim nLinkCount As Long
    Dim strTmp() As String
    ReDim strDownloadLink(0)
    strTmp = Split(Mid(strDecodedSWF, InStr(strDecodedSWF, "url_encoded_fmt_stream_map"), Len(strDecodedSWF)), "url=")
    For nI = 0 To UBound(strTmp)
        If InStr(strTmp(nI), "o-o---preferred---") > 0 Then
            DoEvents
            If InStr(strTmp(nI), ";") > 0 Then
                strDownloadLink(nLinkCount) = Mid(strTmp(nI), 1, InStr(strTmp(nI), ";") - 1) & Replace(Mid(Mid(strTmp(nI), InStr(Replace(strTmp(nI), "&sig=", "&signature="), "&signature=") + 1, Len(strTmp(nI))), 1, InStr(Mid(strTmp(nI), InStr(Replace(strTmp(nI), "&sig=", "&signature="), "&signature=") + 1, Len(strTmp(nI))), "&") - 1), "sig=", "&signature=") & "&quality=" & LCase(ExtractMatch(strTmp(nI), "&quality=([a-zA-Z0-9]*)"))
            Else
                strDownloadLink(nLinkCount) = Replace(Mid(strTmp(nI), 1, InStrRev((strTmp(nI)), ",") - 1), "&sig=", "&signature=")
            End If
            nLinkCount = nLinkCount + 1
            ReDim Preserve strDownloadLink(nLinkCount)
        End If
    Next
End Function
