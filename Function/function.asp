<%
'////////////////////////////////////////////////FSO操作/////////////////////////////////////
'判断文件夹是否存在

Function FolderExits(Folder)
    Folder = Server.MapPath(Folder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FolderExits = True
    Else
        FolderExits = False
    End If
End Function

'判断文件是否存在

Function FileExits(FileName)
    FileName = Server.MapPath(FileName)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function

'创建文件夹

Function CreateFolder(Folder)
    On Error Resume Next
    Folder = Server.MapPath(Folder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    FSO.CreateFolder(Folder)
    If Err>0 Then
        Err.Clear
        CreateFolder = False
    Else
        CreateFolder = True
    End If
End Function

'创建文件

Function CreateFile(FileName, Content)
    On Error Resume Next
    FileName = Server.MapPath(FileName)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
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
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
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

'删除文件夹

Function DeleteFolder(Folder)
    On Error Resume Next
    Folder = Server.MapPath(Folder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
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

'***********************************************
'函数名：ShowfileSize
'作  用：得到文件大小,或文件夹大小
'参  数：str  ----文件或文件夹
'返回值：文件大小X Byte
'***********************************************

Function ShowfileSize(fileorfolder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    Set f = FSO.GetFile(Server.MapPath(fileorfolder))
    ShowfileSize = f.Size
End Function

' 转换字节数为简写形式

Function cSize(tSize)
    If tSize>= 1073741824 Then
        cSize = Int((tSize / 1073741824) * 1000) / 1000 & " GB"
    ElseIf tSize>= 1048576 Then
        cSize = Int((tSize / 1048576) * 1000) / 1000 & " MB"
    ElseIf tSize>= 1024 Then
        cSize = Int((tSize / 1024) * 1000) / 1000 & " KB"
    Else
        cSize = tSize & "B"
    End If
End Function


Function readfile(filename)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(Server.MapPath(filename), 1, True)
    readfile = cnt.ReadAll
End Function

'////////////////////////////////////////////////FSO操作/////////////////////////////////////

'////////////////////////////////////////XMLHTTP/////////////////////////////////////////////

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
        .waitForResponse 1000
        GetBody = .ResponseBody
    End With
    Set xmlhttp = Nothing
End Function

Function BytesToBstr(Body, Cset)
    On Error Resume Next
    Dim objstream
    Set objstream = Server.CreateObject("adodb.stream")
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
    Set xmlhttp = server.CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, false
    xmlhttp.send()
    if xmlhttp.Status<>200 then exit function
    getHTTPimg = xmlhttp.responseBody
    Set xmlhttp = Nothing
    If Err.Number<>0 Then Err.Clear
End Function

Function Save2Local(from, tofile)
    Dim geturl, objStream, imgs
    geturl = Trim(from)
    imgs = gethttpimg(geturl)
    Set objStream = Server.CreateObject("ADODB.Stream")
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

'////////////////////////////////////////XMLHTTP/////////////////////////////////////////////

'////////////////////////////////////////////随机数///////////////////////////////////////////////

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

'/////////////////////////////////////////随机数/////////////////////////////////////////////////


'关键字加链接，要处理的字符，关键字，要添加的链接地址
Function ADDURL(Str,Keywords,URL)
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = Keywords
	ADDURL= regEx.Replace(Str, "<a href="""&URL&"""  id=""cntKeywords"">$1</a>")
	Set regEx = Nothing
End Function 


'检查组件是否被支持及组件版本的子程序

Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Set TestObj = Server.CreateObject (strObj)
    If -2147221005 <> Err Then '感谢网友iAmFisher的宝贵建议
        IsObj = True
        VerObj = TestObj.Version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    Set TestObj = Nothing
End Sub

'z-blog提供，search(内容,关键字)

Function Search(strText, strQuestion)

    Dim s
    Dim i
    Dim j

    s = strText
    i = InStr(1, s, strQuestion, vbTextCompare)
    If i>0 Then
        s = Left(s, i + Len(strQuestion) + 100)
        s = Right(s, Len(strQuestion) + 200)
    Else
        s = ""
    End If

    If s<>"" Then
        i = 1
        Do While InStr(i, s, strQuestion, vbTextCompare)>0
            j = InStr(i, s, strQuestion, vbTextCompare)
            If Len(s) - j - Len(strQuestion)<0 Then
                s = Left(s, j -1) & "<b style='color:#FF6347'>" & strQuestion & "</b>"
                Exit Do
            Else
                s = Left(s, j -1) & "<b style='color:#FF6347'>" & strQuestion & "</b>" & Right(s, Len(s) - j - Len(strQuestion) + 1)
            End If
            i = j + Len("<b style='color:#FF6347'>" & strQuestion & "</b>") -1
            If i>= Len(s) Then Exit Do
        Loop

    End If

    If s = "" Then
        'Search=strText
        Search = Left(delhtml(strText), 200)
    Else
        Search = s
    End If

End Function

'把所有日文删除

Function filteJapanese(sStr)
    Dim oRegExp
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.Pattern = "[\u3040-\u309F\u30A0-\u30FF]"
    filteJapanese = oRegExp.Replace(sStr, "")
    Set oRegExp = Nothing
End Function

'***********************************************
'函数名：HtmlCodeIn;HtmlCodeOut
'作  用：
'文本框写入数据库时--------->HtmlCodeIn,
'从数据库调出到文本框时----->HtmlCodeOut,
'直接在网页上显示则--------->直接调用
'***********************************************
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

'***********************************************
'函数名：DelHTML
'作  用：去除所有HTML标记,主要用在,调用前N个字符
'***********************************************

Function DelHTML(Str)
    '去掉所有HTML标记
    Dim Re, l, t, c, i
    Set Re = New RegExp
    Re.IgnoreCase = True
    Re.Global = True
    're.Pattern = "<(style|script|object|frameset)[\s\S]+?</\1>"
    Re.Pattern = "<(.[^>]*)>"
    DelHTML = Re.Replace(Str, "")
    Set Re = Nothing
End Function

Function DelSpace(Str)
    DelSpace = Trim(Str)
    If DelSpace<>"" Then
        DelSpace = Replace(DelSpace, Chr(10), "")
        DelSpace = Replace(DelSpace, Chr(13), "")
        DelSpace = Replace(DelSpace, Chr(32), "")
        DelSpace = Replace(DelSpace, "&nbsp;", "")
        DelSpace = Replace(DelSpace, "　", "")
    End If
End Function




'ip转换

Function Fn_IP(ip)
    ip = CStr(ip)
    ip1 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip2 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip3 = Left(ip, CInt(InStr(ip, ".") -1))
    ip4 = Mid(ip, CInt(InStr(ip, ".") + 1))
    Fn_IP = CInt(ip1) * 256 * 256 * 256 + CInt(ip2) * 256 * 256 + CInt(ip3) * 256 + CInt(ip4)
End Function


Function Fn_IP1(ip)
    ip = CStr(ip)
    ip1 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip2 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip3 = Left(ip, CInt(InStr(ip, ".") -1))
    Fn_IP1 = CInt(ip1) * 256 * 256 * 256 + CInt(ip2) * 256 * 256 + CInt(ip3) * 256 + 0
End Function

'***********************************************
'函数名：ChangeIP,GetWhere
'作  用：
'ChangeIP--->把IP转换成一个IP所对应的唯一的数字
'GetWhere--->得到给定的IP的地址.
'***********************************************

Function ChangeIP(ip)
    ip = CStr(ip)
    ip1 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip2 = Left(ip, CInt(InStr(ip, ".") -1))
    ip = Mid(ip, CInt(InStr(ip, ".") + 1))
    ip3 = Left(ip, CInt(InStr(ip, ".") -1))
    ip4 = Mid(ip, CInt(InStr(ip, ".") + 1))
    ChangeIP = CInt(ip1) * 256 * 256 * 256 + CInt(ip2) * 256 * 256 + CInt(ip3) * 256 + CInt(ip4)
End Function

Function GetWhere(IP)
    Set rsip = Server.CreateObject("ADODB.Recordset")
    Sql = "Select * from IP_Old UNION Select * from IP_ADD  where IPStart <= "&IP&" And IPEnd >= "&IP&"order by id desc"
    rsip.Open Sql, conn, 1, 1

    If Not rsip.EOF Then
        GetWhere = rsip("country")
    Else
        GetWhere = "未知"
    End If
    rsip.Close
    Set rsip = Nothing
End Function


'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************

Function strlength(Str)
    On Error Resume Next
    Dim winnt_chinese
    winnt_chinese = (Len("中国") = 2)
    If winnt_chinese Then
        Dim l, t, c, i
        l = Len(Str)
        t = l
        For i = 1 To l
            c = Asc(Mid(Str, i, 1))
            If c<0 Then c = c + 65536
            If c>255 Then
                t = t + 1
            End If
        Next
        strlength = t
    Else
        strlength = Len(Str)
    End If
    If Err.Number<>0 Then Err.Clear
End Function



'*************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'*************************************************

Function gotTopic(Str, strlen)
    If Str = "" Or IsNull(Str) Then
        gotTopic = ""
        Exit Function
    End If
    Dim l, t, c, i, m
    m = 0
    Str = Replace(Replace(Replace(Replace(Str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(Str)
    t = 0
    For i = 1 To l
        c = Abs(Asc(Mid(Str, i, 1)))
        If c>255 Then
            t = t + 2
        Else
            t = t + 1
            m = m + 1
        End If
        If t>strlen Then
            If m Mod 2 = 0 Then gotTopic = Left(Str, i) & "…"
            If m Mod 2<>0 Then gotTopic = Left(Str, i + 1) & "…"

            Exit For
        Else
            gotTopic = Str
        End If
    Next
    gotTopic = Replace(Replace(Replace(Replace(gotTopic, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function






'***********************************************
'过程名：showpages1
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************

Sub showpages1(sfilename, totalnumber, maxperpage, ShowTotal, ShowAllPages, strUnit, align)
    Dim n, i, strTemp, strUrl
    If totalnumber Mod maxperpage = 0 Then
        n = totalnumber \ maxperpage
    Else
        n = totalnumber \ maxperpage+1
    End If

    strTemp = "<table align="&align&" ><Form Name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    If ShowTotal = True Then
        strTemp = strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    strUrl = JoinChar(sfilename)
    If CurrentPage<2 Then
        strTemp = strTemp & "首页 上一页&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage -1) & "'>上一页</a>&nbsp;"
    End If

    If n - currentpage<1 Then
        strTemp = strTemp & "下一页 尾页"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
    End If
    strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
    If ShowAllPages = True Then
        strTemp = strTemp & "&nbsp;转到：<Select Name='page' Size='1' onchange='javascript:Submit()'>"
        For i = 1 To n
            strTemp = strTemp & "<Option Value='" & i & "'"
            If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
            strTemp = strTemp & ">第" & i & "页</Option>"
        Next
        strTemp = strTemp & "</Select>"
    End If
    strTemp = strTemp & "</td></tr></Form></table>"
    Response.Write strTemp
End Sub

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'pos=InStr(1,"abcdefg","cd")
'则pos会返回3表示查找到并且位置为第三个字符开始。
'这就是“查找”的实现，而“查找下一个”功能的
'实现就是把当前位置作为起始位置继续查找。
'***********************************************

Function JoinChar(strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?")<Len(strUrl) Then
        If InStr(strUrl, "?")>1 Then
            If InStr(strUrl, "&")<Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function



'***********************************************
'函数名：SetPage;PrintPage
'作  用：分页函数
'一般直接写上就行:setpage(rs,PageCount,20)   PrintPage(PageCount,rs,URL)
'***********************************************

Sub SetPage(rs, PageCount, PageSize)
    If Not rs.EOF Then
        If Not IsEmpty(Request("page")) Then
            If IsNumeric(Request("page")) = True Then
                PageCount = CInt(Request("page"))
            Else
                PageCount = 1
            End If
        Else
            PageCount = 1
        End If
        If Session("PageSize") = "" Or Not IsNumeric(Session("PageSize")) Or PageSize<>Session("PageSize") Then
            Session("PageSize") = PageSize
        End If
        rs.PageSize = Session("PageSize")
        If PageCount>rs.PageCount Or PageCount<= 0 Then
            PageCount = 1
        End If
        If Not rs.EOF Then rs.AbsolutePage = PageCount
    End If
End Sub

Sub PrintPage(PageCount, rs, URL)
    If rs.PageCount > 0 Then
        strurl = "<font size=3>[" & PageCount&"/"&rs.PageCount&"]</font>&nbsp;"
        i = 1
        Do While i<rs.PageCount + 1
            If i<> PageCount Then
                strurl = strurl&"<A href="&URL&"&page="&i&">["&i&"]</A>&nbsp;"
            Else
                strurl = strurl&"<font color=red size=3>["&i&"]</font>&nbsp;"
            End If
            i = i + 1
        Loop
    End If
    Response.Write "　"&strurl
End Sub

Sub PrintPage2(nowPage, URL)
    allpage = Session("PageSize")
    If allPage <= 10 Then
        '总数小于行于10页
        xx = 1
        yy = allPage
    Else
        '总数大于10页
        If nowPage >5 Then
            If nowPage+5 >= allPage Then
                yy = allPage
                xx = allPage -9
            Else
                yy = nowPage+4
                xx = yy -9
            End If
        Else
            xx = 1
            yy = 10
        End If
    End If
    '分页显示
    Str = "<a href="""&URL&""">首页</a>&nbsp;"
    For cc = xx To yy
        If Int(cc)<> Int(nowPage) Then
            Str = Str&"<A href="&URL&"&page="&cc&">"&cc&"</A>&nbsp;"
        Else
            Str = Str&"<font color=red>"&cc&"</font>&nbsp;"
        End If
    Next
    Str = Str&"<a href="""&URL&"&page="&Session("PageSize")&""">尾页</a>"
    Response.Write Str
End Sub




'***********************************************
'函数名：guolvstr
'作  用：过滤掉BBS里的不健康的词句
'Words----->要过滤的字符
'OutStr----->替换成什么字符
'***********************************************

Function guolvstr(Words, OutStr)
    guolvstr = Words
    Const InvaildWords = "fuck|bitch|他妈的|法轮|falundafa|falun|snk.ni8.net|操你妈|三级片|Fa轮功|fa轮功|falun|日你|我日|suck|shit|法轮|我操|李宏治|阴茎|傻B|妈的|操你|干你|日您|屁眼|国民党|台独|卖淫|流氓|999fuck|傻逼|阴道|阳痿|法輪" '需要过滤得字符以“|”隔开
    InvaildWord = Split(InvaildWords, "|")
    For Each abc In InvaildWord
        guolvstr = Replace(guolvstr, abc, OutStr)
    Next
End Function



'***********************************************
'函数名：guolvurl
'作  用：过滤BBS里的网址的.com.org.com.cn等字符及其前边的10个字符为某字符

'***********************************************

Function guolvurl(Words, OutStr)
    guolvurl = Words
    Dim strPattern
    strPattern1 = "(.gov|.cn|.sh|.Name|.ws|.ac|.io|.com|.tw|.idv|.com.cn|.org|.edu)"
    strPattern2 = "(\w{0,10}(.gov|.cn|.sh|.Name|.ws|.ac|.io|.com|.tw|.idv|.com.cn|.org|.edu))\b"
    Dim oRegEx, oMatch
    Set oRegEx = New RegExp
    oRegEx.IgnoreCase = True
    oRegEx.Global = True
    oRegEx.Pattern = strPattern1
    guolvurl = oRegEx.Replace(guolvurl, "$1"&vbCrLf)
    oRegEx.Pattern = strPattern2
    guolvurl = oRegEx.Replace(guolvurl, OutStr)
    Set oRegEx = Nothing
End Function


'格式化Tags
Function FormatTags(Tags)
'可以用正则实现：" |\,+"
Tags=Trim(Tags)
if len(Tags)<=1 then exit function
	Tags=replace(Tags," ","")
	Tags=replace(Tags,"，",",")
	Tags=replace(Tags,"　","")
	Tags=replace(Tags,",,",",")
	if Tags="" then exit function
	if left(Tags,1)="," then Tags=Mid(Tags,2)
	if right(Tags,1)="," then Tags=Mid(Tags,1,len(Tags)-1)
	FormatTags=Tags
End Function


'格式化SQL
Function FormatSQL(FildName,SearchCnt)
	SearchCnt=Trim(replace(SearchCnt,"　"," "))
	if instr(SearchCnt," ")<0 then
		FormatSQL=""
		exit function
	end if
	Set re = New RegExp
	re.Pattern = " +"
	re.Global = True
	re.IgnoreCase = True
	SearchCnt = re.Replace(SearchCnt, "@")
	SearchCnt=split(SearchCnt,"@")
	FildName=split(FildName,",")
	for i_i=0 to ubound(FildName)
			FormatSQL=FormatSQL & " or "
		for i_ii=0 to ubound(SearchCnt)
			FormatSQL=FormatSQL & " and " & FildName(i_i) &" like '%" & SearchCnt(i_ii) &"%'"
		next
	next
	FormatSQL=mid(FormatSQL,4)
	if instr(FormatSQL,"or")>=0 then
		FormatSQL=" and (" & mid(FormatSQL,7)
		FormatSQL=replace(FormatSQL," or  and ",") or (")
		FormatSQL=FormatSQL &")"
	end if
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


'得到包括HTML在内的NO个字符。
Function getHTMLContent(NO, txtContent)
    If Len(txtContent)<= NO Then
        getHTMLContent = txtContent
        Exit Function
    End If

    Dim regEx, Match, Matches, s, E ' 建立变量。
	img=0
    Set regEx = New RegExp ' 建立正则表达式。
    regEx.IgnoreCase = True ' 设置是否区分字符大小写。
    regEx.Global = True ' 设置全局可用性。
    patrn = "<p|<br|<li|<table"
    patrn = Split(patrn, "|")

    For Each p in patrn
        regEx.Pattern = p ' 设置模式。
        txtContent = regEx.Replace(txtContent, "||" & p )
    Next

    txtContent = Split(txtContent, "||")
    For Each E in txtContent
        If Len(delhtml(s))>NO or img>=3 Then Exit For
		if instr(s,"<img")>0 then img=img+1
        s = s + E
    Next
    
        regEx.Pattern = "</?table[\s\S]*?>|</?td[\s\S]*?>|</?tr[\s\S]*?>"
        s = regEx.Replace(s, "" )
    getHTMLContent = s
End Function



'Call SendMail("mailto@163.com", "柳永法", "测试", "<a href=http://www.yongfa365.com/blog/>柳永法test</a>", "local", "sendmail@qq.com")

Function SendMail_temp(MailtoAddress, MailtoName, Subject, MailBody, FromName, MailFrom)
    '函数参数说明（收件人地址,收件人姓名(可选),主题,邮件内容,发件人姓名(可选),发件人地址）
    'on error resume next
    Dim JMail, MailServer, Flag
    Set JMail = Server.CreateObject("JMail.Message")
    If Err Then
        SendMail = False
        Err.Clear
        Exit Function
    End If
    JMail.Charset = "gb2312" '邮件编码
    JMail.silent = True
    JMail.ContentType = "text/html" '邮件正文格式
    MailServer = "smtp.163.com" '用来发送邮件的SMTP服务器
    JMail.MailServerUserName = "用户名" '登录用户名
    JMail.MailServerPassWord = "密码" '登录密码
    JMail.MailDomain = "163.com" '域名（如果用"name@domain.com"这样的用户名登录时，请指明domain.com），可选
    JMail.AddRecipient MailtoAddress, MailtoName '收信人邮箱和收信人名称
    JMail.Subject = Subject '主题
    JMail.AppendHTML(MailBody)
    JMail.FromName = FromName '发信人姓名
    JMail.From = MailFrom '发信人Email
    JMail.Priority = 3 '邮件等级，1为加急，3为普通，5为低级
    Flag = JMail.Send(MailServer)
    If Flag Then
        SendMail = True
    Else
        SendMail = False
    End If
    JMail.Close
    Set JMail = Nothing
    response.Write SendMail
End Function

%>
