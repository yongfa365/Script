Function getContent(patrn, strng)
    Dim regEx, oMatch, Matches '建立变量。
    Set regEx = New RegExp '建立正则表达式。
    regEx.Pattern = patrn'设置模式。
    regEx.IgnoreCase = True '设置是否区分字符大小写。
    regEx.Global = True '设置全局可用性。
    Set Matches = regEx.Execute(strng)'执行搜索。
    For Each oMatch in Matches'遍历匹配集合。
        RetStr = oMatch.SubMatches(0)
        msgbox(RetStr)
    Next
    getContent = RetStr
End Function


Function getHTTPPage(Path)
    t = GetBody(Path)
    getHTTPPage = BytesToBstr(t, "GB2312")
End Function

Function GetBody(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    With xmlhttp
        .Open "Get", url, False, "", ""
        .Send
        If xmlhttp.Status<>200 Then Exit Function
        GetBody = .ResponseBody
    End With
    Set xmlhttp = Nothing
End Function

Function BytesToBstr(Body, Cset)
    On Error Resume Next
    Dim objstream
    Set objstream = CreateObject("adodb.stream")
    objstream.Type = 1
    objstream.Mode = 3
    objstream.Open
    objstream.Write Body
    objstream.Position = 0
    objstream.Type = 2
    objstream.Charset = Cset
    BytesToBstr = objstream.ReadText
    objstream.Close
    Set objstream = Nothing
End Function

Function getHTTPimg(url)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send()
    If xmlhttp.Status<>200 Then Exit Function
    getHTTPimg = xmlhttp.responseBody
    Set xmlhttp = Nothing
    If Err.Number<>0 Then Err.Clear
End Function

Function Save2Local(from, tofile)
    Dim geturl, objStream, imgs
    geturl = Trim(from)
    imgs = gethttpimg(geturl)
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objstream.Write imgs
    objstream.SaveToFile tofile, 2
    objstream.Close()
    Set objstream = Nothing
End Function

Function URLEncode(strURL)
    Dim I
    Dim tempStr
    For I = 1 To Len(strURL)
        If Asc(Mid(strURL, I, 1)) < 0 Then
            tempStr = "%" & Right(CStr(Hex(Asc(Mid(strURL, I, 1)))), 2)
            tempStr = "%" & Left(CStr(Hex(Asc(Mid(strURL, I, 1)))), Len(CStr(Hex(Asc(Mid(strURL, I, 1))))) - 2) & tempStr
            URLEncode = URLEncode & tempStr
        ElseIf (Asc(Mid(strURL, I, 1)) >= 65 And Asc(Mid(strURL, I, 1)) <= 90) Or (Asc(Mid(strURL, I, 1)) >= 97 And Asc(Mid(strURL, I, 1)) <= 122) Then
            URLEncode = URLEncode & Mid(strURL, I, 1)
        Else
            URLEncode = URLEncode & "%" & Hex(Asc(Mid(strURL, I, 1)))
        End If
    Next
End Function

url="http://community.csdn.net/Expert/TopicView1.asp?id=5642424"

            Body = getHTTPPage(url)
            Set re = New RegExp
            re.IgnoreCase = True
            re.Global = True
            
            patrn = "<rank>(.+?)</rank>"
            body = getContent(patrn, body)