On Error Resume Next
truename = InputBox("请输入blog名：")
username = URLEncode(truename)
url = "http://hi.baidu.com/" & username & "/"
Body = getHTTPPage(url)
css = "css/item/(.+\.css)"
cssURL = "http://hi.baidu.com/"&username&"/css/item/"&getContent(css, body)

Str = getHTTPPage(cssURL)
File = Str
CreateFolder truename
CreateFolder truename & "\" & truename
imagess = Split(getImages(Str), "|")
AllArticleLinks = ""
For s = LBound(imagess) To UBound(imagess)
    If InStr(AllArticleLinks, imagess(s))<1 Then
        AllArticleLinks = AllArticleLinks & "||" & imagess(s)
    End If
Next
imagess = Split(AllArticleLinks, "||")

For i = LBound(imagess) To UBound(imagess)
    Save2Local imagess(i), truename & "\" & truename & "\" &i&".gif"
    File = Replace(File, imagess(i), truename & "/" & truename & "/" &i&".gif")
Next

CreateFile truename & "\now.css", File
'----------------------------------------函数区-----------------------------------------------------

Function getImages(Str)
    Set re = New RegExp
    re.Pattern = "url\((.+)?\)"
    re.Global = True
    re.IgnoreCase = True
    Set Contents = re.Execute(Str)
    For Each Match in Contents ' 遍历匹配集合。
        Images = Images + Match.SubMatches(0) + "|"
    Next
    getImages = Mid(Images, 1, Len(Images) -1)
End Function

Function getContent(patrn, strng)
    Dim regEx, oMatch, Matches '建立变量。
    Set regEx = New RegExp '建立正则表达式。
    regEx.Pattern = patrn'设置模式。
    regEx.IgnoreCase = True '设置是否区分字符大小写。
    regEx.Global = True '设置全局可用性。
    Set Matches = regEx.Execute(strng)'执行搜索。
    For Each oMatch in Matches'遍历匹配集合。
        RetStr = oMatch.SubMatches(0)
        Exit For
    Next
    getContent = RetStr
End Function

Function RegExpTest(patrn, strng)
    Dim regEx, Match, Matches '建立变量。
    Set regEx = New RegExp '建立正则表达式。
    regEx.Pattern = patrn'设置模式。
    regEx.IgnoreCase = True '设置是否区分字符大小写。
    regEx.Global = True '设置全局可用性。
    Set Matches = regEx.Execute(strng)'执行搜索。
    For Each Match in Matches'遍历匹配集合。
        RetStr = RetStr & "||" & Match.Value
    Next
    RegExpTest = RetStr
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

Function CreateFile(FileName, Content)
    On Error Resume Next
    FileName = MapPath(FileName)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fd = FSO.CreateTextFile(FileName, True)
    fd.WriteLine Content
    If Err>0 Then
        Err.Clear
        CreateFile = False
    Else
        CreateFile = True
    End If
End Function

Function CreateFolder(Folder)
    On Error Resume Next
    Folder = MapPath(Folder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CreateFolder(Folder)
    If Err>0 Then
        Err.Clear
        CreateFolder = False
    Else
        CreateFolder = True
    End If
End Function
