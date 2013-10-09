
'Wscript.Echo GetChildsID(Conn, AllIDs, "GetCategory", "Href", "ParentHref", "")

'GetChildID(���ݿ����Ӷ���,����ձ���,TableName,IDField,ParentIDField,ParentIDText)
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




'<---------------------------------------FSO����--------------------------------------->
'�ж��ļ����Ƿ����
Function FolderExits(Folder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FolderExits = True
    Else
        FolderExits = False
    End If
End Function

'�ж��ļ��Ƿ����
Function FileExits(FileName)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function

'�����ļ���
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

'ɾ���ļ���
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

'�����ļ�
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

'ɾ���ļ�
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

'��ȡ�ļ�
Function ReadFile(filename)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(filename, 1, True)
    readfile = cnt.ReadAll
End Function

'��ʾ�ļ���С
Function ShowFileSize(FileOrFolder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set f = FSO.GetFile(FileOrFolder)
    ShowfileSize = f.Size
End Function

'�ı���ʾ��С��λ
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

'<---------------------------------------/ FSO����------------------------------------->



'<---------------------------------------Adodb.Stream����--------------------------------------->
'��ָ�������ȡ�ļ�
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

'��ָ������洢�ļ�
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
'<---------------------------------------/ Adodb.Stream����------------------------------------->


'<---------------------------------------XMLHTTP------------------------------------->

Function getHTTPPage2(Path)
    t = GetBody(Path)
    getHTTPPage = BytesToBstr(t, "GB2312")
End Function

Function getHTTPPage(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "Get", url, False
    xmlhttp.setRequestHeader "Referer", "http://www.google.com/" '�趨ҳ����Դ 
	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Googlebot/2.1;++http://www.google.com/bot.html)" '�趨����� 
	
'	xmlhttp.setRequestHeader "Cookie", sSetCookie '�趨Cookie 
'	xmlhttp.setRequestHeader "Referer", sReferer '�趨ҳ����Դ 
'	xmlhttp.setRequestHeader "Accept-Language", "zh-cn" '�趨���� 
'	xmlhttp.setRequestHeader "Content-Length",Len(sData) '�趨���ݳ��� 
'	xmlhttp.setRequestHeader "CONTENT-Type",sCONTENT '�趨������������ 
'	xmlhttp.setRequestHeader "Accept-Encoding", sEncoding '�趨gzipѹ�� 
'	xmlhttp.setRequestHeader "Accept", "gb2312" '�ĵ����� 

'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(Windows;+U;+Windows+NT+5.2;+zh-CN;+rv:1.9.0.1)+Gecko/2008070208+Firefox/3.0.1" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Googlebot/2.1;++http://www.google.com/bot.html)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Baiduspider+(+http://www.baidu.com/search/spider.htm)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Yahoo!+Slurp+China;+http://misc.yahoo.com.cn/help.html)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;+Yahoo!+Slurp;+http://help.yahoo.com/help/us/ysearch/slurp)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "msnbot/1.1+(+http://search.msn.com/msnbot.htm)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0+(compatible;YodaoBot-Reader/1.0;http://www.yodao.com/help/webmaster/spider/;1+subscriber;)" '�趨����� 
'	xmlhttp.setRequestHeader "User-Agent", "Sogou+web+spider/4.0(+http://www.sogou.com/docs/help/webmasters.htm#07)" '�趨����� 

    xmlhttp.Send
    If xmlhttp.Status<>200 Then Exit Function
    GetBody = xmlhttp.ResponseBody
    '��ͷ�ļ��￴����
    GetCodePage = getContents("charset=[""']*([^""']+)", xmlhttp.getResponseHeader("Content-Type") , True)(0)
    '�ڷ��ص��ַ����￴����Ȼ���������룬����Ӱ������ȡ�����
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

'<---------------------------------------�����------------------------------------->

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

'<---------------------------------------/ �����------------------------------------->

'<---------------------------------------Error--------------------------------------->
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
'<---------------------------------------/ Error------------------------------------->

'<---------------------------------------Common--------------------------------------->
'�ؼ��ּ����ӣ�Ҫ������ַ����ؼ��֣�Ҫ��ӵ����ӵ�ַ
Function ADDURL(Str,Keywords,URL)
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = Keywords
	ADDURL= regEx.Replace(Str, "<a href="""&URL&"""  id=""cntKeywords"">$1</a>")
	Set regEx = Nothing
End Function 

'SQL��ע�뺯�������÷���������Ҫ��ע��ĵط��滻��ǰ��request("XXXX")ΪSafeRequest("XXXX")   
'www.yongfa365.com   
  
Function SafeRequest(ParaValue)
    ParaValue = Trim(Request(ParaValue))
    '������ʽ����
    Set re = New RegExp
    '��ֹʹ�õ�ע���ַ�
    re.Pattern = "\'|Select|Update|Delete|insert|Count|drop table|truncate|Asc|Mid|char|xp_cmdshell|exec master|net localgroup administrators|And|net user|Or"
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(ParaValue)
    RegExpTest = Matches.count
    'ע�봦��
    If RegExpTest >0 Then
        Response.Write "<script language='javascript'>alert('���ɵ�SQLע������!');window.history.go(-1);</script>"
        response.End
    Else
        SafeRequest = ParaValue
    End If
End Function

'�������Ƿ�֧�ּ�����汾���ӳ���

Sub ObjTest(strObj)
    On Error Resume Next
    IsObj = False
    VerObj = ""
    Set TestObj = CreateObject (strObj)
    If -2147221005 <> Err Then '��л����iAmFisher�ı�����
        IsObj = True
        VerObj = TestObj.Version
        If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
    End If
    Set TestObj = Nothing
End Sub

'����������ɾ��

Function KillJapanese(sStr)
    Dim re
    Set re = New RegExp
    re.Global = True
    re.Pattern = "[\u3040-\u309F\u30A0-\u30FF]"
    KillJapanese = re.Replace(sStr, "")
    Set re = Nothing
End Function

'ת����������ΪHTML��ʽ
function ChangeJapanese(l1)
		dim I1,I2,i,l2:l2=l1
		I1=array("��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��","��")
		I2=array("460","462","463","450","466","468","470","472","474","476","478","480","482","485","487","489","496","497","499","500","502","503","505","506","508","509","532")
		for i=0 to 26
			l2=replace(l2,I1(i),"&#12"&I2(i)&";")
		next
		ChangeJapanese=l2
end function

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


'***********************************************
'��������DelHTML
'��  �ã�ȥ������HTML���,��Ҫ����,����ǰN���ַ�
'***********************************************

Function DelHTML(Str)
    'ȥ������HTML���
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
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
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
'��������GetLeftChar
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       StrLen ----��ȡ����
'����ֵ����ȡ����ַ���
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

'�ַ�������
function StringFilter(Str,OutStr)
	set re = new RegExp
	re.IgnoreCase=False
	re.Global=True
	re.Pattern="fuck|bitch|�����|����|falundafa|falun|snk.ni8.net|������|����Ƭ|Fa�ֹ�|fa�ֹ�|falun|����|����|suck|shit|����|�Ҳ�|�����|����|ɵB|���|����|����|����|ƨ��|����|̨��|����|��å|999fuck|ɵ��|����|����|��݆" '��Ҫ���˵��ַ��ԡ�|������
	StringFilter=re.Replace(StringFilter,OutStr)
end Function

'������ʽд���滻
Function reReplace(Str, restrS, restrD)
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = restrS
    reReplace = re.Replace(Str, restrD)
End Function


'��ʽ��Tags
Function FormatTags(Tags)
	set re = new RegExp
	re.IgnoreCase=False
	re.Global=True
	re.Pattern="��|;|��| |	|��"
	Tags=re.replace(Tags,",")
	re.Pattern=",{2,}"
	Tags=re.replace(Tags,",")
	re.Pattern="^[\s,]*(\S+?)[\s,]*$"
	Tags=re.replace(Tags,"$1")
	FormatTags=Tags
End Function

'���������û��ƥ������
Function RegExpTest(patrn, strng)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    RegExpTest = Matches.count
End Function

'aaa = GetContents("<a href=(.*?) alt=(.*?)>(.*?)</a>", "<a href=http://www.yongfa365.com alt=����������վ>������</a>  <a href=http://www.y.com alt=>��</a>  <a href=http://www.365.com alt=��>365</a>" , True)
'For i = 0 To UBound(aaa) step 3
'    WScript.Echo aaa(i)
'    WScript.Echo aaa(i + 1)
'    WScript.Echo aaa(i + 2)
'Next
'WScript.Echo GetContents("a(.+?)b", "a23234b ab a67896896b sadfasdfb" , True)(0)

'�õ�ƥ������ݣ�����������,ǰһ��Ҫ��IsArray()�����ԣ���治��
'������ʽ���ַ������Ƿ񷵻�����ֵ
'�������飬UBoundû��ƥ�䷵��-1,��������>=0

Function GetContents(Pattern, Str , BoolGetSubMatches)
    SplitStr = "��������˭"
    OKStr = ""
    
    Set re = New RegExp
    re.Pattern = Pattern
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(Str)
    
    '���ûƥ��ɹ����򷵻ؿ�����
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


'ɾ�������ظ���
Function DelArrayDup(Strs)
    SplitStr = "������"
    NewStrs = SplitStr
    For Each Str in Strs
        If(InStr(NewStrs, SplitStr & Str & SplitStr) = 0) Then NewStrs = NewStrs + Str + SplitStr
    Next
    DelArrayDup = Split(Mid(NewStrs, Len(SplitStr) + 1, Len(NewStrs) - Len(SplitStr) * 2), SplitStr)
End Function

'instring str2��Ҫ�����ַ�,��str2�ǲ�����str1���
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
