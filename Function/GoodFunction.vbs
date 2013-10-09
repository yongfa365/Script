
'Wscript.Echo GetChildsID(Conn, AllIDs, "GetCategory", "Href", "ParentHref", "")

'GetChildID(数据库连接对象,任意空变量,TableName,IDField,ParentIDField,ParentIDText)
Function GetChildsID(Conn, AllIDs, TableName, IDField, ParentIDField, ParentIDText)
    sql = "select " & IDField & "," & ParentIDField & " from " & TableName & " where " & ParentIDField & " = '" & ParentIDText & "'"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        AllIDs = AllIDs & "," & rs(IDField)
        GetChildsID Conn, AllIDs, TableName, IDField, ParentIDField, rs(IDField)
        rs.movenext
    Loop
    GetChildsID = Mid(AllIDs, 2)
End Function

'Wscript.Echo GetChildsIDByCode(Conn,"Admin_City","Code","")

Function GetChildsIDByCode(Conn, TableName, CodeFields, NowCode)
    sql = "select " & CodeFields & " from " & TableName & " where " & CodeFields & " like'" & NowCode & "__'"
    Set rs = Conn.Execute(sql)
    Do While Not rs.EOF
        AllIDs = AllIDs & "," & rs(CodeFields)
        sql2 = "select " & CodeFields & " from " & TableName & " where " & CodeFields & " like'" & rs(CodeFields) & "_%'"
        Set rs2 = Conn.Execute(sql2)
        Do While Not rs2.EOF
            AllIDs = AllIDs & "," & rs2(CodeFields)
            rs2.movenext
        Loop
        rs.movenext
    Loop
    GetChildsIDByCode = Mid(AllIDs,2)
End Function




'<---------------------------------------FSO操作--------------------------------------->
'判断文件夹是否存在
Function FolderExits(Folder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FolderExits = True
    Else
        FolderExits = False
    End If
End Function

'判断文件是否存在
Function FileExits(FileName)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function

'创建文件夹
Function CreateFolder(Folder)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CreateFolder(Folder)
    If Err>0 Then
        Err.Clear
        CreateFolder = False
    Else
        CreateFolder = True
    End If
End Function

'删除文件夹
Function DeleteFolder(Folder)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FSO.DeleteFolder Folder, True
    End If
    If Err>0 Then
        Err.Clear
        DeleteFolder = False
    Else
        DeleteFolder = True
    End If
End Function

'创建文件
Function CreateFile(FileName, Content)
    On Error Resume Next
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

'删除文件
Function DeleteFile(FileName)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FSO.DeleteFile FileName, True
    End If
    If Err>0 Then
        Err.Clear
        DeleteFile = False
    Else
        DeleteFile = True
    End If
End Function

'读取文件
Function ReadFile(filename)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(filename, 1, True)
    readfile = cnt.ReadAll
End Function

'显示文件大小
Function ShowFileSize(FileOrFolder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set f = FSO.GetFile(FileOrFolder)
    ShowfileSize = f.Size
End Function

'改变显示大小单位
Function ChangeSize(sizes)
    If sizes<1024 Then
        ChangeSize = sizes & " Byte"
    End If
    If sizes>1024 Then
        sizes = (sizes / 1024)
        ChangeSize = FormatNumber(sizes, 2) & " KB"
    End If
    If sizes>1024 Then
        sizes = (sizes / 1024)
        ChangeSize = FormatNumber(sizes, 2) & " MB"
    End If
    If sizes>1024 Then
        sizes = (sizes / 1024)
        ChangeSize = FormatNumber(sizes, 2) & " GB"
    End If
End Function

'<---------------------------------------/ FSO操作------------------------------------->



'<---------------------------------------Adodb.Stream操作--------------------------------------->
'按指定编码读取文件
Function ReadFromFile(FileUrl, CharSet)
    Dim Str
    Set stm = CreateObject("Adodb.Stream")
    stm.Type = 2
    stm.mode = 3
    stm.charset = CharSet
    stm.Open
    stm.loadfromfile FileUrl
    Str = stm.readtext
    stm.Close
    Set stm = Nothing
    ReadFromFile = Str
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
'<---------------------------------------/ Adodb.Stream操作------------------------------------->


'<---------------------------------------XMLHTTP------------------------------------->

Function getHTTPPage2(Path)
    t = GetBody(Path)
    getHTTPPage = BytesToBstr(t, "GB2312")
End Function

Function getHTTPPage(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "Get", url, False
    xmlhttp.setRequestHeader "Referer", "http://www.google.com/" '设定页面来源 
	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Googlebot/2.1;++http://www.google.com/bot.html)" '设定浏览器 
	
'	xmlhttp.setRequestHeader "Cookie", sSetCookie '设定Cookie 
'	xmlhttp.setRequestHeader "Referer", sReferer '设定页面来源 
'	xmlhttp.setRequestHeader "Accept-Language", "zh-cn" '设定语言 
'	xmlhttp.setRequestHeader "Content-Length",Len(sData) '设定数据长度 
'	xmlhttp.setRequestHeader "CONTENT-Type",sCONTENT '设定接受数据类型 
'	xmlhttp.setRequestHeader "Accept-Encoding", sEncoding '设定gzip压缩 
'	xmlhttp.setRequestHeader "Accept", "gb2312" '文档类型 

'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(Windows;+U;+Windows+NT+5.2;+zh-CN;+rv:1.9.0.1)+Gecko/2008070208+Firefox/3.0.1" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Googlebot/2.1;++http://www.google.com/bot.html)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Baiduspider+(+http://www.baidu.com/search/spider.htm)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Yahoo!+Slurp+China;+http://misc.yahoo.com.cn/help.html)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Yahoo!+Slurp;+http://help.yahoo.com/help/us/ysearch/slurp)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "msnbot/1.1+(+http://search.msn.com/msnbot.htm)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;YodaoBot-Reader/1.0;http://www.yodao.com/help/webmaster/spider/;1+subscriber;)" '设定浏览器 
'	xmlhttp.setRequestHeader "User-Agent", "Sogou+web+spider/4.0(+http://www.sogou.com/docs/help/webmasters.htm#07)" '设定浏览器 

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

Function GetBody(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    With xmlhttp
        .Open "Get", url, False
        .Send
        .waitForResponse 1000
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
    xmlhttp.Open "GET", url, false
    xmlhttp.send()
    if xmlhttp.Status<>200 then exit function
    getHTTPimg = xmlhttp.responseBody
    Set xmlhttp = Nothing
    If Err>0 Then Err.Clear
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

'Call Save2Local("http://www.baidu.com/img/logo.gif", "F:\baidulogo.gif")
'Call Save2Local("http://flash.jninfo.net/images/banner.swf", "F:\banner.swf")
'Call Save2Local("http://www.yongfa365.com/", "F:\1.html")
'response.write getHTTPPage("http://www.yongfa365.com/")

'<---------------------------------------/ XMLHTTP------------------------------------->

'<---------------------------------------随机数------------------------------------->

Function NumRand(n) '生成n位随机数字
    For i = 1 To n
        Randomize
        temp = CInt(9 * Rnd)
        temp = temp + 48
        NumRand = NumRand & Chr(temp)
    Next
End Function


Function LCharRand(n) '生成n位随机小写字母
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        temp = temp + 97
        LCharRand = LCharRand & Chr(temp)
    Next
End Function


Function UCharRand(n) '生成n位随机大写字母
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        temp = temp + 65
        UCharRand = UCharRand & Chr(temp)
    Next
End Function


Function allRand(n) '生成n位随机数字字母子组合
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        If temp Mod 2 = 0 Then
            temp = temp + 97
        ElseIf temp < 9 Then
            temp = temp + 48
        Else
            temp = temp + 65
        End If
        allRand = allRand & Chr(temp)
    Next
End Function

Function LallRand(n) '生成n位随机数字字母子组合
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        If temp Mod 2 = 0 Then
            temp = temp + 97
        ElseIf temp < 9 Then
            temp = temp + 48
        Else
            temp = temp + 97
        End If
        LallRand = LallRand & Chr(temp)
    Next
End Function

'1234567890转化为A-J的字母

Function ChangeCode(id)
    For i = 1 To Len(Trim(id))
        ChangeCode = ChangeCode + Chr(Mid(id, i, 1) + 65)
    Next

End Function

'<---------------------------------------/ 随机数------------------------------------->

'<---------------------------------------Error--------------------------------------->
'****************************************************
'过程名：ErrorMsg;SuccessMsg,ReturnNOMsg
'作  用：显示错误,正确提示信息,直接跳转到一个地方。
'参  数：str;str1,str2字符型
'****************************************************

Function ErrorMsg(Str)
    Response.Write "<script language='javascript'>alert('"&Str&"');window.history.go(-1);</script>"
    Response.End
End Function


Function SuccessMsg(str1, str2)
    Response.Write "<script language='javascript'>alert('"&str1&"');window.location.href='"&str2&"';</script>"
    Response.End
End Function


Function RetnrnNoMsg(Str)
    Response.Write "<script language='javascript'>window.location.href='"&Str&"';</script>"
    Response.End
End Function
'<---------------------------------------/ Error------------------------------------->

'<---------------------------------------Common--------------------------------------->
'关键字加链接，要处理的字符，关键字，要添加的链接地址
Function ADDURL(Str,Keywords,URL)
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = Keywords
	ADDURL= regEx.Replace(Str, "<a href="""&URL&"""  id=""cntKeywords"">$1</a>")
	Set regEx = Nothing
End Function 

'SQL防注入函数，调用方法，在需要防注入的地方替换以前的request("XXXX")为SafeRequest("XXXX")   
'www.yongfa365.com   
  
Function SafeRequest(ParaValue)
    ParaValue = Trim(Request(ParaValue))
    '正则表达式过滤
    Set re = New RegExp
    '禁止使用的注入字符
    re.Pattern = "\'|Select|Update|Delete|insert|Count|drop table|truncate|Asc|Mid|char|xp_cmdshell|exec master|net localgroup administrators|And|net user|Or"
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(ParaValue)
    RegExpTest = Matches.count
    '注入处理
    If RegExpTest >0 Then
        Response.Write "<script language='javascript'>alert('可疑的SQL注入请求!');window.history.go(-1);</script>"
        response.End
    Else
        SafeRequest = ParaValue
    End If
End Function

'检查组件是否被支持及组件版本的子程序

Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Set TestObj = CreateObject (strObj)
    If -2147221005 <> Err Then '感谢网友iAmFisher的宝贵建议
        IsObj = True
        VerObj = TestObj.Version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    Set TestObj = Nothing
End Sub

'把所有日文删除

Function KillJapanese(sStr)
    Dim re
    Set re = New RegExp
    re.Global = True
    re.Pattern = "[\u3040-\u309F\u30A0-\u30FF]"
    KillJapanese = re.Replace(sStr, "")
    Set re = Nothing
End Function

'转换所有日文为HTML格式
function ChangeJapanese(l1)
		dim I1,I2,i,l2:l2=l1
		I1=array("ガ","ギ","グ","ア","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","パ","ビ","ピ","ブ","プ","ベ","ペ","ボ","ポ","ヴ")
		I2=array("460","462","463","450","466","468","470","472","474","476","478","480","482","485","487","489","496","497","499","500","502","503","505","506","508","509","532")
		for i=0 to 26
			l2=replace(l2,I1(i),"&#12"&I2(i)&";")
		next
		ChangeJapanese=l2
end function

'***********************************************
'函数名：HtmlCodeIn;HtmlCodeOut
'作  用：
'文本框写入数据库时--------->HtmlCodeIn,
'从数据库调出到文本框时----->HtmlCodeOut,
'直接在网页上显示则--------->直接调用
'***********************************************
Function HtmlCodeIn(fString)
	If Trim(fString)<>"" Then
		fString = Replace(fString, "&", "&amp;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, """", "&quot;")
		'fString = Replace(fString, " ", "&nbsp;")
		fString = Replace(fString, vbcrlf, "<br />")
	End If
	HtmlCodeIn = fString
End Function

Function HtmlCodeOut(fString)
	If Trim(fString)<>"" Then
		fString = Replace(fString, "&lt;", "<")
		fString = Replace(fString, "&gt;",">" )
		fString = Replace(fString, "&amp;","&" )
		fString = Replace(fString, "&quot;","""")
		'fString = Replace(fString, "&nbsp;"," ")
		fString = Replace(fString, "<br />",vbcrlf )
	End If
	HtmlCodeOut = fString
End Function


'***********************************************
'函数名：DelHTML
'作  用：去除所有HTML标记,主要用在,调用前N个字符
'***********************************************

Function DelHTML(Str)
    '去掉所有HTML标记
    Dim Re, l, t, c, i
    Set Re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    're.Pattern = "<(style|script|object|frameset)[\s\S]+?</\1>"
    re.Pattern = "<(.[^>]*)>"
    DelHTML = re.Replace(Str, "")
    Set Re = Nothing
End Function


'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************

Function StringLength(Str)
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "[^\x00-\xFF]"
    StringLength = len(re.Replace(Str, "ok"))
    Set Re = Nothing
End Function

'*************************************************
'函数名：GetLeftChar
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       StrLen ----截取长度
'返回值：截取后的字符串
'*************************************************
Function GetLeftChar(Str, StrLen)
	Str=Trim(Str)
    If Str = "" Or IsNull(Str) Then
        GetLeftChar = ""
        Exit Function
    End If
    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "[^\x00-\xFF]"
    
	if StrLen>=len(re.replace(Str,"aa")) then 
		GetLeftChar=Str
		exit Function
	End If
	
    t = 0
    For i = 1 To Len(Str)
        If re.test(mid(Str,i,1)) Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t>StrLen Then
            GetLeftChar = Left(Str, i) & ".."
            Exit For
        End If
    Next
    set re=Nothing
End Function

'字符串过滤
function StringFilter(Str,OutStr)
	set re = new RegExp
	re.IgnoreCase=False
	re.Global=True
	re.Pattern="fuck|bitch|他妈的|法轮|falundafa|falun|snk.ni8.net|操你妈|三级片|Fa轮功|fa轮功|falun|日你|我日|suck|shit|法轮|我操|李宏治|阴茎|傻B|妈的|操你|干你|日您|屁眼|国民党|台独|卖淫|流氓|999fuck|傻逼|阴道|阳痿|法輪" '需要过滤得字符以“|”隔开
	StringFilter=re.Replace(StringFilter,OutStr)
end Function

'正则表达式写的替换
Function reReplace(Str, restrS, restrD)
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = restrS
    reReplace = re.Replace(Str, restrD)
End Function


'格式化Tags
Function FormatTags(Tags)
	set re = new RegExp
	re.IgnoreCase=False
	re.Global=True
	re.Pattern="，|;|；| |	|　"
	Tags=re.replace(Tags,",")
	re.Pattern=",{2,}"
	Tags=re.replace(Tags,",")
	re.Pattern="^[\s,]*(\S+?)[\s,]*$"
	Tags=re.replace(Tags,"$1")
	FormatTags=Tags
End Function

'正则测试有没有匹配内容
Function RegExpTest(patrn, strng)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    RegExpTest = Matches.count
End Function

'aaa = GetContents("<a href=(.*?) alt=(.*?)>(.*?)</a>", "<a href=http://www.yongfa365.com alt=柳永法的网站>柳永法</a>  <a href=http://www.y.com alt=>永</a>  <a href=http://www.365.com alt=法>365</a>" , True)
'For i = 0 To UBound(aaa) step 3
'    WScript.Echo aaa(i)
'    WScript.Echo aaa(i + 1)
'    WScript.Echo aaa(i + 2)
'Next
'WScript.Echo GetContents("a(.+?)b", "a23234b ab a67896896b sadfasdfb" , True)(0)

'得到匹配的内容，并返回数组,前一版要用IsArray()来测试，这版不用
'正则表达式，字符串，是否返回引用值
'返回数组，UBound没有匹配返回-1,其它返回>=0

Function GetContents(Pattern, Str , BoolGetSubMatches)
    SplitStr = "柳永法是谁"
    OKStr = ""
    
    Set re = New RegExp
    re.Pattern = Pattern
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(Str)
    
    '如果没匹配成功，则返回空数组
    If Matches.Count = 0 Then
        GetContents = Split("", SplitStr)
        Exit Function
    End If
    
    If BoolGetSubMatches Then
        For Each oMatch in Matches
            If oMatch.SubMatches.Count>0 Then
                For Each math in oMatch.SubMatches
                    OKStr = OKStr & math & SplitStr
                Next
            End If
        Next
    Else
        For Each oMatch in Matches
            If oMatch.Value<>"" Then OKStr = OKStr & oMatch.Value & SplitStr
        Next
    End If
    
    If Len(OKStr)> Len(SplitStr) Then
        GetContents = Split(Left(OKStr, Len(OKStr) - Len(SplitStr)), SplitStr)
    Else
        GetContents = Split("", SplitStr)
    End If
End Function


'删除数组重复项
Function DelArrayDup(Strs)
    SplitStr = "柳永法"
    NewStrs = SplitStr
    For Each Str in Strs
        If(InStr(NewStrs, SplitStr & Str & SplitStr) = 0) Then NewStrs = NewStrs + Str + SplitStr
    Next
    DelArrayDup = Split(Mid(NewStrs, Len(SplitStr) + 1, Len(NewStrs) - Len(SplitStr) * 2), SplitStr)
End Function

'instring str2是要检测的字符,看str2是不是在str1里边
Function ISinstr(str1, str2)
    str1 = Trim(str1)
    If str1 = "" Or IsNull(str1) Then
        ISinstr = False
    Else
        mystr = Split(str1, ",")
        For i = LBound(mystr) To UBound(mystr)
            If mystr(i)<>"" Then
                If instr(mystr(i),str2) > 0 Then
                    ISinstr = True
                    Exit Function
                End If
            End If
        Next
        ISinstr = False
    End If
End Function


'<---------------------------------------/ Common------------------------------------->
