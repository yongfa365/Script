on error resume next
SpaceName="onlytest001"
username="onlytest001"
password="onlytest001"
loginstr = "username=" + username + "&password=" + password
loginUrl = "http://passport.baidu.com/?login"
Set xmlhttp = Createobject("Microsoft.XMLHTTP")
xmlhttp.Open "POST",loginUrl,False
xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
xmlhttp.Send loginstr

postURL="http://hi.baidu.com/"&SpaceName&"/commit"
dbpath = "Admin5.com.mdb"
connstr = "provider=microsoft.jet.oledb.4.0;data source=" & dbpath
Set conn = CreateObject("adodb.connection")
conn.Open connstr
Set rs = CreateObject("Adodb.Recordset")
sql= "select * from Articles where Category='幽默网文' order by id desc"
'sql= "select top 1000 * from Articles"
rs.Open sql, conn, 1, 3
for i=1 to rs.recordcount

		title=URLEncode(rs("txtTitle"))&i
		content=URLEncode(rs("txtContent"))
		Cat=URLEncode("默认分类")
		WriteContent = "spBlogTitle="&title&"&spBlogCatName="&Cat&"&spIsCmtAllow=1&spBlogPower=0&ct=1&cm=1&spBlogText="&content
'			title=urlencoding(rs("txtTitle")&i)
'		content=urlencoding(rs("txtContent"))
'		Cat=urlencoding("默认分类")
		WriteContent = "spBlogTitle="&title&"&spBlogCatName="&Cat&"&spIsCmtAllow=1&spBlogPower=0&ct=1&cm=1&spBlogText="&content
'	
		Set xmlhttp = Createobject("Microsoft.XMLHTTP")
		xmlhttp.Open "POST",postURL,False
		xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
		xmlhttp.setRequestHeader "Referer", postURL '设定页面来源 
		xmlhttp.Send WriteContent
WriteToFile i&".txt", WriteContent, "gb2312" 
WriteToFile i&".html", BytesToBstr(xmlhttp.ResponseBody,"gb2312"), "gb2312" 
WScript.Echo i
WScript.Sleep 1000*10
rs.movenext
next
'
'If xmlhttp.Status=200 Then GetBody=xmlhttp.ResponseBody
'msgbox OK
'writetofile "baidu.html",BytesToBstr(GetBody,"gb2312"),"gb2312"

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

'按指定编码存储文件
Function WriteToFile (FileUrl, Str, CharSet)
    Set stm = CreateObject("Adodb.Stream")
    stm.Type = 2
    stm.mode = 3
    stm.charset = CharSet
    stm.Open
    stm.WriteText Str
    stm.SaveToFile FileUrl, 2
    stm.flush
    stm.Close
    Set stm = Nothing
End Function

Function URLEncode(strURL)
    Dim I
    Dim tempStr
    For I = 1 To Len(strURL)
    ASC_Temp=Asc(Mid(strURL, I, 1))
        If ASC_Temp < 0 Then
            tempStr = "%" & Right(CStr(Hex(ASC_Temp)), 2)
            tempStr = "%" & Left(CStr(Hex(ASC_Temp)), Len(CStr(Hex(ASC_Temp))) - 2) & tempStr
            URLEncode = URLEncode & tempStr
        ElseIf (ASC_Temp >= 65 And ASC_Temp <= 90) Or (ASC_Temp >= 97 And ASC_Temp <= 122) Or (ASC_Temp >= 48 And ASC_Temp <= 57)  Then
            URLEncode = URLEncode & Mid(strURL, I, 1)
        ElseIf ASC_Temp=10 or ASC_Temp=13 then
        	'10 13 9 vbcrlf 
        Else
            URLEncode = URLEncode & "%" & Hex(ASC_Temp)
        End If
    Next
End Function

Function Urlcode(InpStr) '字符串转换为UrlCode编码
    Dim InpAsc, I
    For I = 1 To Len(InpStr) '取单个字符处理
        'msgbox Mid(InpStr, I, 1) & " // " & Asc(Mid(InpStr, I, 1))
        InpAsc = Asc(Mid(InpStr, I, 1)) '单个字符的ASC码
        If ((InpAsc < 58) And (InpAsc > 47)) Or ((InpAsc < 91) And (InpAsc > 64)) Or ((InpAsc < 123) And (InpAsc > 96)) Then
            Urlcode = Urlcode & Chr(InpAsc) '如果是数字或字母原字符不变
        Else
            Urlcode = Urlcode & "%" & Mid(Trim(Hex(InpAsc)), 1, 2) & "%" & Mid(Trim(Hex(InpAsc)), 3, 2)
        End If
    Next
End Function

Function urlencoding(vstrin)
    Dim i, strreturn
    strreturn = ""
    For i = 1 To Len(vstrin)
        thischr = Mid(vstrin, i, 1)
        If Abs(Asc(thischr)) < &hff Then
            strreturn = strreturn & thischr
        Else
            innercode = Asc(thischr)
            If innercode < 0 Then
                innercode = innercode + &h10000
            End If
            hight8 = (innercode And &hff00) \ &hff
            low8 = innercode And &hff
            strreturn = strreturn & "%" & Hex(hight8) & "%" & Hex(low8)
        End If
    Next
    urlencoding = strreturn
End Function
