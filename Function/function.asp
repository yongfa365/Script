<%
'////////////////////////////////////////////////FSO����/////////////////////////////////////
'�ж��ļ����Ƿ����

Function FolderExits(Folder)
    Folder = Server.MapPath(Folder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FolderExits = True
    Else
        FolderExits = False
    End If
End Function

'�ж��ļ��Ƿ����

Function FileExits(FileName)
    FileName = Server.MapPath(FileName)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function

'�����ļ���

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

'�����ļ�

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

'ɾ���ļ�

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

'ɾ���ļ���

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
'��������ShowfileSize
'��  �ã��õ��ļ���С,���ļ��д�С
'��  ����str  ----�ļ����ļ���
'����ֵ���ļ���СX Byte
'***********************************************

Function ShowfileSize(fileorfolder)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    Set f = FSO.GetFile(Server.MapPath(fileorfolder))
    ShowfileSize = f.Size
End Function

' ת���ֽ���Ϊ��д��ʽ

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

'////////////////////////////////////////////////FSO����/////////////////////////////////////

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

'////////////////////////////////////////////�����///////////////////////////////////////////////

Function NumRand(n) '����nλ�������
    For i = 1 To n
        Randomize
        temp = CInt(9 * Rnd)
        temp = temp + 48
        NumRand = NumRand & Chr(temp)
    Next
End Function


Function LCharRand(n) '����nλ���Сд��ĸ
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        temp = temp + 97
        LCharRand = LCharRand & Chr(temp)
    Next
End Function


Function UCharRand(n) '����nλ�����д��ĸ
    For i = 1 To n
        Randomize
        temp = CInt(25 * Rnd)
        temp = temp + 65
        UCharRand = UCharRand & Chr(temp)
    Next
End Function


Function allRand(n) '����nλ���������ĸ�����
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

Function LallRand(n) '����nλ���������ĸ�����
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

'1234567890ת��ΪA-J����ĸ

Function ChangeCode(id)
    For i = 1 To Len(Trim(id))
        ChangeCode = ChangeCode + Chr(Mid(id, i, 1) + 65)
    Next

End Function

'/////////////////////////////////////////�����/////////////////////////////////////////////////


'�ؼ��ּ����ӣ�Ҫ������ַ����ؼ��֣�Ҫ��ӵ����ӵ�ַ
Function ADDURL(Str,Keywords,URL)
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = Keywords
	ADDURL= regEx.Replace(Str, "<a href="""&URL&"""  id=""cntKeywords"">$1</a>")
	Set regEx = Nothing
End Function 


'�������Ƿ�֧�ּ�����汾���ӳ���

Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Set TestObj = Server.CreateObject (strObj)
    If -2147221005 <> Err Then '��л����iAmFisher�ı�����
        IsObj = True
        VerObj = TestObj.Version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    Set TestObj = Nothing
End Sub

'z-blog�ṩ��search(����,�ؼ���)

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

'����������ɾ��

Function filteJapanese(sStr)
    Dim oRegExp
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.Pattern = "[\u3040-\u309F\u30A0-\u30FF]"
    filteJapanese = oRegExp.Replace(sStr, "")
    Set oRegExp = Nothing
End Function

'***********************************************
'��������HtmlCodeIn;HtmlCodeOut
'��  �ã�
'�ı���д�����ݿ�ʱ--------->HtmlCodeIn,
'�����ݿ�������ı���ʱ----->HtmlCodeOut,
'ֱ������ҳ����ʾ��--------->ֱ�ӵ���
'***********************************************
'***********************************************
'��������HtmlCodeIn;HtmlCodeOut
'��  �ã�
'�ı���д�����ݿ�ʱ--------->HtmlCodeIn,
'�����ݿ�������ı���ʱ----->HtmlCodeOut,
'ֱ������ҳ����ʾ��--------->ֱ�ӵ���
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
'��������ErrorMsg;SuccessMsg,ReturnNOMsg
'��  �ã���ʾ����,��ȷ��ʾ��Ϣ,ֱ����ת��һ���ط���
'��  ����str;str1,str2�ַ���
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
'��������DelHTML
'��  �ã�ȥ������HTML���,��Ҫ����,����ǰN���ַ�
'***********************************************

Function DelHTML(Str)
    'ȥ������HTML���
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
        DelSpace = Replace(DelSpace, "��", "")
    End If
End Function




'ipת��

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
'��������ChangeIP,GetWhere
'��  �ã�
'ChangeIP--->��IPת����һ��IP����Ӧ��Ψһ������
'GetWhere--->�õ�������IP�ĵ�ַ.
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
        GetWhere = "δ֪"
    End If
    rsip.Close
    Set rsip = Nothing
End Function


'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************

Function strlength(Str)
    On Error Resume Next
    Dim winnt_chinese
    winnt_chinese = (Len("�й�") = 2)
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
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
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
            If m Mod 2 = 0 Then gotTopic = Left(Str, i) & "��"
            If m Mod 2<>0 Then gotTopic = Left(Str, i + 1) & "��"

            Exit For
        Else
            gotTopic = Str
        End If
    Next
    gotTopic = Replace(Replace(Replace(Replace(gotTopic, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function






'***********************************************
'��������showpages1
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
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
        strTemp = strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    strUrl = JoinChar(sfilename)
    If CurrentPage<2 Then
        strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage -1) & "'>��һҳ</a>&nbsp;"
    End If

    If n - currentpage<1 Then
        strTemp = strTemp & "��һҳ βҳ"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
    End If
    strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp = strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
    If ShowAllPages = True Then
        strTemp = strTemp & "&nbsp;ת����<Select Name='page' Size='1' onchange='javascript:Submit()'>"
        For i = 1 To n
            strTemp = strTemp & "<Option Value='" & i & "'"
            If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
            strTemp = strTemp & ">��" & i & "ҳ</Option>"
        Next
        strTemp = strTemp & "</Select>"
    End If
    strTemp = strTemp & "</td></tr></Form></table>"
    Response.Write strTemp
End Sub

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'pos=InStr(1,"abcdefg","cd")
'��pos�᷵��3��ʾ���ҵ�����λ��Ϊ�������ַ���ʼ��
'����ǡ����ҡ���ʵ�֣�����������һ�������ܵ�
'ʵ�־��ǰѵ�ǰλ����Ϊ��ʼλ�ü������ҡ�
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
'��������SetPage;PrintPage
'��  �ã���ҳ����
'һ��ֱ��д�Ͼ���:setpage(rs,PageCount,20)   PrintPage(PageCount,rs,URL)
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
    Response.Write "��"&strurl
End Sub

Sub PrintPage2(nowPage, URL)
    allpage = Session("PageSize")
    If allPage <= 10 Then
        '����С������10ҳ
        xx = 1
        yy = allPage
    Else
        '��������10ҳ
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
    '��ҳ��ʾ
    Str = "<a href="""&URL&""">��ҳ</a>&nbsp;"
    For cc = xx To yy
        If Int(cc)<> Int(nowPage) Then
            Str = Str&"<A href="&URL&"&page="&cc&">"&cc&"</A>&nbsp;"
        Else
            Str = Str&"<font color=red>"&cc&"</font>&nbsp;"
        End If
    Next
    Str = Str&"<a href="""&URL&"&page="&Session("PageSize")&""">βҳ</a>"
    Response.Write Str
End Sub




'***********************************************
'��������guolvstr
'��  �ã����˵�BBS��Ĳ������Ĵʾ�
'Words----->Ҫ���˵��ַ�
'OutStr----->�滻��ʲô�ַ�
'***********************************************

Function guolvstr(Words, OutStr)
    guolvstr = Words
    Const InvaildWords = "fuck|bitch|�����|����|falundafa|falun|snk.ni8.net|������|����Ƭ|Fa�ֹ�|fa�ֹ�|falun|����|����|suck|shit|����|�Ҳ�|�����|����|ɵB|���|����|����|����|ƨ��|����|̨��|����|��å|999fuck|ɵ��|����|����|��݆" '��Ҫ���˵��ַ��ԡ�|������
    InvaildWord = Split(InvaildWords, "|")
    For Each abc In InvaildWord
        guolvstr = Replace(guolvstr, abc, OutStr)
    Next
End Function



'***********************************************
'��������guolvurl
'��  �ã�����BBS�����ַ��.com.org.com.cn���ַ�����ǰ�ߵ�10���ַ�Ϊĳ�ַ�

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


'��ʽ��Tags
Function FormatTags(Tags)
'����������ʵ�֣�" |\,+"
Tags=Trim(Tags)
if len(Tags)<=1 then exit function
	Tags=replace(Tags," ","")
	Tags=replace(Tags,"��",",")
	Tags=replace(Tags,"��","")
	Tags=replace(Tags,",,",",")
	if Tags="" then exit function
	if left(Tags,1)="," then Tags=Mid(Tags,2)
	if right(Tags,1)="," then Tags=Mid(Tags,1,len(Tags)-1)
	FormatTags=Tags
End Function


'��ʽ��SQL
Function FormatSQL(FildName,SearchCnt)
	SearchCnt=Trim(replace(SearchCnt,"��"," "))
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
    Dim regEx, oMatch, Matches '����������
    Set regEx = New RegExp '����������ʽ��
    regEx.Pattern = patrn'����ģʽ��
    regEx.IgnoreCase = True '�����Ƿ������ַ���Сд��
    regEx.Global = True '����ȫ�ֿ����ԡ�
    Set Matches = regEx.Execute(strng)'ִ��������
    For Each oMatch in Matches'����ƥ�伯�ϡ�
        RetStr = oMatch.SubMatches(0)
        Exit For
    Next
    getContent = RetStr
End Function


'�õ�����HTML���ڵ�NO���ַ���
Function getHTMLContent(NO, txtContent)
    If Len(txtContent)<= NO Then
        getHTMLContent = txtContent
        Exit Function
    End If

    Dim regEx, Match, Matches, s, E ' ����������
	img=0
    Set regEx = New RegExp ' ����������ʽ��
    regEx.IgnoreCase = True ' �����Ƿ������ַ���Сд��
    regEx.Global = True ' ����ȫ�ֿ����ԡ�
    patrn = "<p|<br|<li|<table"
    patrn = Split(patrn, "|")

    For Each p in patrn
        regEx.Pattern = p ' ����ģʽ��
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



'Call SendMail("mailto@163.com", "������", "����", "<a href=http://www.yongfa365.com/blog/>������test</a>", "local", "sendmail@qq.com")

Function SendMail_temp(MailtoAddress, MailtoName, Subject, MailBody, FromName, MailFrom)
    '��������˵�����ռ��˵�ַ,�ռ�������(��ѡ),����,�ʼ�����,����������(��ѡ),�����˵�ַ��
    'on error resume next
    Dim JMail, MailServer, Flag
    Set JMail = Server.CreateObject("JMail.Message")
    If Err Then
        SendMail = False
        Err.Clear
        Exit Function
    End If
    JMail.Charset = "gb2312" '�ʼ�����
    JMail.silent = True
    JMail.ContentType = "text/html" '�ʼ����ĸ�ʽ
    MailServer = "smtp.163.com" '���������ʼ���SMTP������
    JMail.MailServerUserName = "�û���" '��¼�û���
    JMail.MailServerPassWord = "����" '��¼����
    JMail.MailDomain = "163.com" '�����������"name@domain.com"�������û�����¼ʱ����ָ��domain.com������ѡ
    JMail.AddRecipient MailtoAddress, MailtoName '���������������������
    JMail.Subject = Subject '����
    JMail.AppendHTML(MailBody)
    JMail.FromName = FromName '����������
    JMail.From = MailFrom '������Email
    JMail.Priority = 3 '�ʼ��ȼ���1Ϊ�Ӽ���3Ϊ��ͨ��5Ϊ�ͼ�
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
