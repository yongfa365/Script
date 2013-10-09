msgbox "start"
S_URL="http://www.yongfa365.com/articles.xml"
D_File="C:\Program Files\Thunder Network\Thunder\Program\adtask.xml"
body=getHTTPPage(S_URL)

patrn = "<title>.{9}(.+?).{3}</title>"
titles = getContents(patrn, body)
patrn = "<link>(.+?)</link>"
links = getContents(patrn, body)

for i=1 to 10
	content = content & vbcrlf  & "<item index=""100"&i&""" date=""2007-10-01 00:00:00"" url="""&titles(i)&""" desc="""" deltime=""2010-10-24 00:00:00"" clickurl="""&links(i)&""" target=""0"" clienttype=""10"" threshold=""30""></item>"
next

CreateFile D_File,reReplace(readfile(D_File), "<status_ads>([\s\S]+)</status_ads>", "<status_ads>"&content&"</status_ads>")

msgbox "end"
'///////////////////////////////////////////////////
Function readfile(filename)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(filename, 1, True)
    readfile = cnt.ReadAll
End Function

Function FileExits(FileName)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function

Function getContents(patrn, strng)
    Dim regEx, oMatch, Matches '建立变量。
    Set regEx = New RegExp '建立正则表达式。
    regEx.Pattern = patrn'设置模式。
    regEx.IgnoreCase = True '设置是否区分字符大小写。
    regEx.Global = True '设置全局可用性。
    Set Matches = regEx.Execute(strng)'执行搜索。
    ok=0
    For i=1 to Matches.count'遍历匹配集合。
    if i>20 then exit for
    if Matches(i).value<>"" then RetStr = RetStr & Matches(i).SubMatches(0) & "@@@"
    Next

    getContents = split(RetStr,"@@@")
End Function


Function CreateFile(FileName, Content)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fd = FSO.CreateTextFile(FileName, True)
    fd.WriteLine Content
End Function

Function getHTTPPage(Path)
    t = GetBody(Path)
    getHTTPPage = BytesToBstr(t, "utf-8")
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

Function reReplace(Str, restrS, restrD)
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = restrS
    reReplace = re.Replace(Str, restrD)
End Function