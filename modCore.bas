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
    If strWebHTML = "" Then
        SeperateSWF = "Hey! No data recieved! Check your network connection! Or retry!": Exit Function
    End If
    strDecodedSWF = Mid(Mid(strWebHTML, InStr(strWebHTML, "var swf ="), InStr(strWebHTML, ".innerHTML = swf;") - InStr(strWebHTML, "var swf =")), InStr(Mid(strWebHTML, InStr(strWebHTML, "var swf ="), InStr(strWebHTML, ".innerHTML = swf;") - InStr(strWebHTML, "var swf =")), "url_encoded_fmt_stream_map"))
    
    frmMain.Text1.Text = UTF8_UrlDecode(UTF8_UrlDecode(UTF8_UrlDecode(strDecodedSWF)))
End Function

Public Function LoadVideoInfo(ByVal strVideoLink As String, ByVal txtVideoID As TextBox, ByVal txtVideoTitle As TextBox, ByVal txtVideoViews As TextBox, ByVal txtVideoLength As TextBox, ByVal txtVideoUploader As TextBox, ByVal txtVideoChannel As TextBox, ByVal picVideoScreenShotCon As PictureBox, ByVal picVideoScreenShotView As PictureBox)
    LoadVideoID strVideoLink, txtVideoID
    LoadVideoTitle txtVideoTitle
    LoadVideoViews txtVideoViews
    LoadVideoLength txtVideoLength
    LoadAuthorInfo txtVideoUploader, txtVideoChannel
    LoadPicScreenShot picVideoScreenShotCon, picVideoScreenShotView
End Function

Private Sub LoadVideoID(ByVal strVideoLink As String, ByVal txtVideoID As TextBox)
    txtVideoID.Text = ExtractMatch(strVideoLink, "v=([A-Za-z0-9-_]+)")
End Sub

Private Sub LoadVideoTitle(ByVal txtVideoTitle As TextBox)
    txtVideoTitle.Text = Mid(strWebHTML, InStr(strWebHTML, "<title>") + Len("<title>"), InStr(strWebHTML, "</title>") - InStr(strWebHTML, "<title>") - 7)
End Sub

Private Sub LoadVideoViews(ByVal txtVideoViews As TextBox)
    txtVideoViews.Text = Mid(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), InStr(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), "<strong>") + Len("<strong>"), InStr(Mid(strWebHTML, InStr(strWebHTML, "<span class=""watch-view-count"">") + Len("<span class=""watch-view-count"">"), InStr(InStr(strWebHTML, "<span class=""watch-view-count"">"), strWebHTML, "</span>") - InStr(strWebHTML, "<span class=""watch-view-count"">") - Len("<span class=""watch-view-count"">")), "</strong>") - Len("</strong>       "))
    
End Sub

Private Sub LoadVideoLength(ByVal txtVideoLength As TextBox)
    txtVideoLength.Text = ConvertToTime(Val(Mid(strWebHTML, InStr(strWebHTML, ";length_seconds=") + Len(";length_seconds="))))
End Sub

Private Sub LoadAuthorInfo(ByVal txtVideoUploader As TextBox, ByVal txtVideoChannel As TextBox)
    Dim strTmp As String
    strTmp = Trim(Mid(strWebHTML, InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype="""), InStr(InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype="""), strWebHTML, "</span>") - InStr(strWebHTML, "<span itemprop=""author"" itemscope itemtype=""")))
    txtVideoChannel.Text = Left(Mid(strTmp, InStr(strTmp, "<link itemprop=""url"" href=""") + Len("<link itemprop=""url"" href=""")), Len(Mid(strTmp, InStr(strTmp, "<link itemprop=""url"" href=""") + Len("<link itemprop=""url"" href="""))) - 2)
    txtVideoUploader.Text = Replace(txtVideoChannel.Text, "http://www.youtube.com/user/", "")
End Sub

Private Sub LoadPicScreenShot(ByVal picVideoScreenShot As PictureBox, ByVal picVideoScreenShotView As PictureBox)
    
    Set picVideoScreenShot.Picture = LoadPicture(Mid(Mid(strWebHTML, InStr(strWebHTML, "<meta property=""og:image"" content=") + Len("<meta property=""og:image"" content=")), 2, InStr(Mid(strWebHTML, InStr(strWebHTML, "<meta property=""og:image"" content=") + Len("<meta property=""og:image"" content=")), """>") - 2))
    FitPictureToBox picVideoScreenShot, picVideoScreenShotView
    'picVideoScreenShotView.Picture = picVideoScreenShot.Picture
End Sub
