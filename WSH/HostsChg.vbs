WScript.Echo "开始更新"
t=timer
File="C:\WINDOWS\System32\Drivers\Etc\hosts"
Content=getHTTPPage("http://www.yongfa365.com/hosts.txt")

set fso=createobject("scripting.filesystemobject") 
set f=fso.getfile(File) 
f.attributes=32
CreateFile File,Content
f.attributes=33
WScript.Echo "更新成功!共用时：" & (timer-t) & "秒"
'----------------------------------------函数区-----------------------------------------------------

Function CreateFile(FileName, Content)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fd = FSO.CreateTextFile(FileName, True)
    fd.WriteLine Content
End Function

'得到匹配的内容，并以数据形式显示
'表达式，字符串，是否返回引用值
'msgbox getContents("a(.+?)b", "a23234b ab a67896896b sadfasdfb" ,True)(0)

Function getContents(patrn, strng , yinyong)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    If yinyong Then
        For i = 0 To Matches.Count -1
            If Matches(i).Value<>"" Then RetStr = RetStr & Matches(i).SubMatches(0) & "柳永法"
        Next
    Else
        For Each oMatch in Matches
            If oMatch.Value<>"" Then RetStr = RetStr & oMatch.Value & "柳永法"
        Next
    End If
    getContents = Split(RetStr, "柳永法")
End Function

Function getHTTPPage(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "Get", url, False
    xmlhttp.Send
    If xmlhttp.Status<>200 Then Exit Function
    GetBody = xmlhttp.ResponseBody
    '在头文件里看编码
    GetCodePage = getContents("charset=[""']*([^""']+)", xmlhttp.getResponseHeader("Content-Type") , True)(0)
    '在返回的字符串里看，虽然中文是乱码，但不影响我们取其编码
    If Len(GetCodePage)<3 Then GetCodePage = getContents("charset=[""']*([^""']+)", xmlhttp.ResponseText , True)(0)
    If Len(GetCodePage)<3 Then GetCodePage = "gb2312"
    Set xmlhttp = Nothing
    getHTTPPage = BytesToBstr(GetBody, GetCodePage)
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
