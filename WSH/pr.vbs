on error resume next
    connstr = "provider=microsoft.jet.oledb.4.0;data source=username.mdb"
    Set conn = CreateObject("adodb.connection")
    conn.Open connstr
    Set rsuser = CreateObject("Adodb.Recordset")
    sql = "select * from username where pr=100"
	rsuser.open sql,conn,1,3
	do while not rsuser.eof 
		URL = "http://hi.baidu.com/"&rsuser("urlname")
		sURL = "http://www.google.com/search?client=navclient-auto&ch=" & CalculateChecksum(URL) & "&features=Rank&q=info:" & URL
		body = getHTTPPage(sURL)
		pr=int(Mid(body, 10, 11))
		if len(pr)<>0 then
			rsuser("pr")=pr
			rsuser("prOKTime")=now
			rsuser.update
		end if
		wscript.echo  "pr="&right("0"&pr,1) & "-->"& rsuser("id")  & "-->"& URL&vbcrlf
		wscript.sleep 15*1000
		rsuser.movenext
	loop

'////////////////////////////////////////////////////////////////////////////////////////////////////
Const GOOGLE_MAGIC = &HE6359A60

Function sl(ByVal x, ByVal n)
    If n = 0 Then
        sl = x
    Else
        Dim k
        k = CLng(2 ^ (32 - n - 1))
        Dim d
        d = x And (k - 1)
        Dim c
        c = d * CLng(2 ^ n)
        If x And k Then
            c = c Or &H80000000
        End If
        sl = c
    End If
End Function

'//from www.aspxuexi.com

Function sr(ByVal x, ByVal n)
    If n = 0 Then
        sr = x
    Else
        Dim y
        y = x And &H7FFFFFFF
        Dim z
        If n = 32 - 1 Then
            z = 0
        Else
            z = y \ CLng(2 ^ n)
        End If
        If y <> x Then
            z = z Or CLng(2 ^ (32 - n - 1))
        End If
        sr = z
    End If
End Function

Function zeroFill(ByVal a, ByVal b)
    Dim x
    If (&H80000000 And a) Then
        x = sr(a, 1)
        x = x And (Not &H80000000)
        x = x Or &H40000000
        x = sr(x, b - 1)
    Else
        x = sr(a, b)
    End If
    zeroFill = x
End Function

Private Function uadd(ByVal L1, ByVal L2)
    Dim L11, L12, L21, L22, L31, L32
    L11 = L1 And &HFFFFFF
    L12 = (L1 And &H7F000000) \ &H1000000
    If L1 < 0 Then L12 = L12 Or &H80
    L21 = L2 And &HFFFFFF
    L22 = (L2 And &H7F000000) \ &H1000000
    If L2 < 0 Then L22 = L22 Or &H80
    L32 = L12 + L22
    L31 = L11 + L21
    If (L31 And &H1000000) Then L32 = L32 + 1
    uadd = (L31 And &HFFFFFF) + (L32 And &H7F) * &H1000000
    If L32 And &H80 Then uadd = uadd Or &H80000000
End Function

Private Function usub(ByVal L1, ByVal L2)
    Dim L11, L12, L21, L22, L31, L32
    L11 = L1 And &HFFFFFF
    L12 = (L1 And &H7F000000) \ &H1000000
    If L1 < 0 Then L12 = L12 Or &H80
    L21 = L2 And &HFFFFFF
    L22 = (L2 And &H7F000000) \ &H1000000
    If L2 < 0 Then L22 = L22 Or &H80
    L32 = L12 - L22
    L31 = L11 - L21
    If L31 < 0 Then
        L32 = L32 - 1
        L31 = L31 + &H1000000
    End If
    usub = L31 + (L32 And &H7F) * &H1000000
    If L32 And &H80 Then usub = usub Or &H80000000
End Function

Function mix(ByVal ia, ByVal ib, ByVal ic)
    Dim a, b, c
    a = ia
    b = ib
    c = ic
    
    a = usub(a, b)
    a = usub(a, c)
    a = a Xor zeroFill(c, 13)
    
    b = usub(b, c)
    b = usub(b, a)
    b = b Xor sl(a, 8)
    
    c = usub(c, a)
    c = usub(c, b)
    c = c Xor zeroFill(b, 13)
    
    a = usub(a, b)
    a = usub(a, c)
    a = a Xor zeroFill(c, 12)
    
    b = usub(b, c)
    b = usub(b, a)
    b = b Xor sl(a, 16)
    
    c = usub(c, a)
    c = usub(c, b)
    c = c Xor zeroFill(b, 5)
    
    a = usub(a, b)
    a = usub(a, c)
    a = a Xor zeroFill(c, 3)
    
    b = usub(b, c)
    b = usub(b, a)
    b = b Xor sl(a, 10)
    
    c = usub(c, a)
    c = usub(c, b)
    c = c Xor zeroFill(b, 15)
    
    Dim ret(3)
    
    ret(0) = a
    ret(1) = b
    ret(2) = c
    
    mix = ret
End Function

Function gc(ByVal s, ByVal i)
    gc = Asc(Mid(s, i + 1, 1))
End Function

Function GoogleCH(ByVal sUrl)
    Dim iLength, a, b, c, k, iLen, m
    iLength = Len(sUrl)
    
    a = &H9E3779B9
    b = &H9E3779B9
    c = GOOGLE_MAGIC
    k = 0
    
    iLen = iLength
    Do While iLen >= 12
        a = uadd(a, (uadd(gc(sUrl, k + 0), uadd(sl(gc(sUrl, k + 1), 8), uadd(sl(gc(sUrl, k + 2), 16), sl(gc(sUrl, k + 3), 24))))))
        b = uadd(b, (uadd(gc(sUrl, k + 4), uadd(sl(gc(sUrl, k + 5), 8), uadd(sl(gc(sUrl, k + 6), 16), sl(gc(sUrl, k + 7), 24))))))
        c = uadd(c, (uadd(gc(sUrl, k + 8), uadd(sl(gc(sUrl, k + 9), 8), uadd(sl(gc(sUrl, k + 10), 16), sl(gc(sUrl, k + 11), 24))))))
        
        m = mix(a, b, c)
        
        a = m(0)
        b = m(1)
        c = m(2)
        
        k = k + 12
        
        iLen = iLen - 12
    Loop
    
    c = uadd(c, iLength)
    
    Select Case iLen ' all the case statements fall through
        Case 11
            c = uadd(c, sl(gc(sUrl, k + 10), 24))
            c = uadd(c, sl(gc(sUrl, k + 9), 16))
            c = uadd(c, sl(gc(sUrl, k + 8), 8))
            b = uadd(b, sl(gc(sUrl, k + 7), 24))
            b = uadd(b, sl(gc(sUrl, k + 6), 16))
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 10
            c = uadd(c, sl(gc(sUrl, k + 9), 16))
            c = uadd(c, sl(gc(sUrl, k + 8), 8))
            b = uadd(b, sl(gc(sUrl, k + 7), 24))
            b = uadd(b, sl(gc(sUrl, k + 6), 16))
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 9
            c = uadd(c, sl(gc(sUrl, k + 8), 8))
            b = uadd(b, sl(gc(sUrl, k + 7), 24))
            b = uadd(b, sl(gc(sUrl, k + 6), 16))
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 8
            b = uadd(b, sl(gc(sUrl, k + 7), 24))
            b = uadd(b, sl(gc(sUrl, k + 6), 16))
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 7
            b = uadd(b, sl(gc(sUrl, k + 6), 16))
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 6
            b = uadd(b, sl(gc(sUrl, k + 5), 8))
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 5
            b = uadd(b, gc(sUrl, k + 4))
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 4
            a = uadd(a, sl(gc(sUrl, k + 3), 24))
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 3
            a = uadd(a, sl(gc(sUrl, k + 2), 16))
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 2
            
            '//form http://www.aspxuexi.com
            a = uadd(a, sl(gc(sUrl, k + 1), 8))
            a = uadd(a, gc(sUrl, k + 0))
        Case 1
            a = uadd(a, gc(sUrl, k + 0))
    End Select
    
    m = mix(a, b, c)
    
    GoogleCH = m(2)
End Function

Function CalculateChecksum(sUrl)
    CalculateChecksum = "6" & CStr(GoogleCH("info:" & sUrl))
End Function

'//////////////////////////////////////////


Function getHTTPPage(Path)
    t = GetBody(Path)
    getHTTPPage = BytesToBstr(t, "GB2312")
End Function

Function GetBody(url)
    On Error Resume Next
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    With xmlhttp
        .Open "Get", url, False
'		.setRequestHeader "Cookie", sSetCookie '设定Cookie 
'		.setRequestHeader "Referer", sReferer '设定页面来源 
		.setRequestHeader "Accept-Language", "zh-cn" '设定语言 
'		.setRequestHeader "Content-Length",Len(sData) '设定数据长度 
'		.setRequestHeader "CONTENT-Type",sCONTENT '设定接受数据类型 
		.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" '设定浏览器 
'		.setRequestHeader "Accept-Encoding", sEncoding '设定gzip压缩 
'		.setRequestHeader "Accept", "gb2312" '文档类型 
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
