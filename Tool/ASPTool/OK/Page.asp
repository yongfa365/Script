<%
'支持多风格变换的ASP分页类 说明文件

'ACCESS 20万数据时用7m,好点机器上应该更快，在SQL上30万多感觉没有延迟
'Set Page1 = New Page
'Page1.CurrentPage = Trim(Request("Page"))
'Page1.PageNum = 20
'Page1.IsDeBug = True
'Page1.Url = NowFile & "?dev=www.yongfa365.com" & SearchURL & "&"
'Page1.Exec TableName, "id", 1, "" , " 1=1 " & sql, conn
''''Exec("TableName","ID",1,",ID1 desc,ID2 asc,ID3 desc"," ID<>0 and ID>1",Conn)
''''Exec(tblName, fldName, OrderType,OtherOrderBy, strWhere, Conn)
'set rs=Page1.PageRs


'For i = 1 To Page1.PageNum
'If rs.EOF Then Exit For
'	内容列表
'rs.MoveNext
'Next


'Page1.Temp="{N1} {N2} {N3} {N4} || 共:{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条"
'Page1.show(1)
'response.write("<p><p><p>")
'Page1.Temp="[{N1}] [{N2}] [{N3}] [{N4}]"
'Page1.show(1)
'response.write("<p><p><p>")
'Page1.Temp="{N1} {N2} {L}{N}{L/}{N3}{N4}"
'Page1.show(2)
'response.write("<p><p><p>")
'Page1.Temp="{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4}"
'Page1.show(3)
'response.write("<p><p><p>")
'Page1.Temp="{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4} 共{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条"
'Page1.show(4)
'response.write("<p><p><p>")
%>



<%
'支持多风格变换的ASP分页类

Class Page
    Private CurrPage
    Private PageN
    Private UrlStr
    Private TempStr
    Private ErrInfo
    Private IsErr
    Private TotalRecord
    Private SQL_Count
    Private SQL_Page
    Private ConnObj
    Private TotalPage
    Public PageRs
    Public IsDebug



    Private TempA(11)
    Private TempB(8)
    '------------------------------------------------------------

    Private Sub Class_Initialize()
        CurrPage = 1'//默认显示当前页为第一页
        PageN = 10'//默认每页显示10条数据
        UrlStr = "#"
        TempStr = "{N1} {N2} {N3} {N4} || 共:{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条"
        ErrInfo = "错误提示:"
        IsErr = False
        IsDebug=False
    End Sub


    Private Sub Class_Terminate()
        If IsObject(PageRs) Then
            PageRs.Close
            Set PageRs = Nothing
        End If
        Erase TempA
        Erase TempB
    End Sub

    '----------------------------------------------------------
    '//获取当前页码

    Public Property Let CurrentPage(Val)
        CurrPage = Val
    End Property


    Public Property Get CurrentPage()
        CurrentPage = CurrPage
    End Property

    '//获取每页显示条数

    Public Property Let PageNum(Val)
        PageN = Val
    End Property


    Public Property Get PageNum()
        PageNum = PageN
    End Property

    '//获取URL

    Public Property Let Url(Val)
        UrlStr = Val
    End Property


    Public Property Get Url()
        Url = UrlStr
    End Property

    '//获取模板

    Public Property Let Temp(Val)
        TempStr = Val
    End Property


    Public Property Get Temp()
        Temp = TempStr
    End Property

    '------------------------------------------------------------
    'Exec("TableName","ID",1,",ID1 desc,ID2 asc,ID3 desc"," ID<>0 and ID>1",Conn)
    Public Sub Exec(tblName, fldName, OrderType,OtherOrderBy, strWhere, Conn)
        Dim strTemp, strSQL, strOrder
        ConnObj = Conn

        '根据排序方式生成相关代码
        If OrderType = 0 Then
            strTemp = ">(select max([" & fldName & "])"
            strOrder = " order by [" & fldName & "] asc"
        Else
            strTemp = "<(select min([" & fldName & "])"
            strOrder = " order by [" & fldName & "] desc"
        End If
        strOrder =strOrder & OtherOrderBy

        If IsNumeric(CurrPage) Then
            CurrPage = CLng(CurrPage)
            If CurrPage<1 Then CurrPage = 1
        Else
            CurrPage = 1
        End If


        '若是第1页则无须复杂的语句
        If CurrPage = 1 Then
            strTemp = ""
            If strWhere<>"" Then
                strTmp = " where " + strWhere
            End If
            strSQL = "select top " & PageNum & " * from [" & tblName & "]" & strTmp & strOrder
        Else '若不是第1页，构造SQL语句
            strSQL = "select top " & PageNum & " * from [" & tblName & "] where [" & fldName & "]" & strTemp & _
                     " from (select top " & (CurrPage -1) * PageNum & " [" & fldName & "] from [" & tblName & "]"
            If strWhere<>"" Then
                strSQL = strSQL & " where " & strWhere
            End If
            strSQL = strSQL & strOrder & ") as tblTemp)"
            If strWhere<>"" Then
                strSQL = strSQL & " And " & strWhere
            End If
            strSQL = strSQL & strOrder
        End If

        SQL_Page = strSQL '返回SQL语句
        SQL_Count = "select count(*) from [" & tblName & "] where " & strWhere
        
        If IsDebug Then Response.Write SQL_Page
        
        Set PageRs = Server.CreateObject("ADODB.RecordSet")
        PageRs.Open SQL_Page, ConnObj, 0, 1

        If Err.Number<>0 Then
            Err.Clear
            PageRs.Close
            Set PageRs = Nothing
            ErrInfo = ErrInfo&"建立或打开记录集错误..."
            IsErr = True
            Response.Write ErrInfo
            Response.End
        End If

    End Sub

    Private Sub AllTotal()
        PageRs.Close
        PageRs.Open SQL_Count, ConnObj, 0, 1
        TotalRecord = PageRs(0)

        TempI = TotalRecord Mod PageNum
        If TempI>0 Then
            TotalPage = TotalRecord \ PageNum + 1
        Else
            TotalPage = TotalRecord \ PageNum
        End If
    End Sub


    Private Sub Init()
        AllTotal() '这个在数据量很大时如果要统计数据量，会很郁闷，所以放在这，只有调用显示分页函数时才会计算
        
        'Private TempA(10)
        TempA(1) = "{N1}" '//首页
        TempA(2) = "{N2}"'//上一页
        TempA(3) = "{N3}"'//下一页
        TempA(4) = "{N4}"'//尾页
        TempA(5) = "{N5}"'//当前页码
        TempA(6) = "{N6}"'//页码总数
        TempA(7) = "{N7}"'//每页条数
        TempA(8) = "{N8}"'//文章总数
        TempA(9) = "{L}"'//循环标签开始
        TempA(10) = "{N}"'//循环内单标签：页码
        TempA(11) = "{L/}"'//循环标签结束
        'Private TempB(8)
        TempB(1) = "首页"
        TempB(2) = "上一页"
        TempB(3) = "下一页"
        TempB(4) = "尾页"
        TempB(5) = CurrPage'//当前页码
        TempB(6) = TotalPage'//页码总数
        TempB(7) = PageN'//每页条数
        TempB(8) = TotalRecord'//文章总数


    End Sub


    Public Sub Show(Style)
        If IsErr = True Then
            Response.Write ErrInfo
            Exit Sub
        End If

        Call Init()
        Select Case Style
            Case 1
                Response.Write StyleA()
            Case 2
                Response.Write StyleB()
            Case 3
                Response.Write StyleC()
            Case 4
                Response.Write StyleD()
            Case Else
                ErrInfo = ErrInfo&"不存在当前样式..."
                Response.Write ErrInfo
        End Select
    End Sub


    Public Function ShowStyle(Style)
        If IsErr = True Then
            ShowStyle = ErrInfo
            Exit Function
        End If

        Call Init()
        Select Case Style
            Case 1
                ShowStyle = StyleA()
            Case 2
                ShowStyle = StyleB()
            Case 3
                ShowStyle = StyleC()
            Case 4
                ShowStyle = StyleD()
            Case Else
                ErrInfo = ErrInfo&"不存在当前样式..."
                ShowStyle = ErrInfo
        End Select
    End Function

    Private Function StyleA()
        '首页 上一页 下一页 尾页  本页为第1/20页，共20页，每页10条，文章总数200条
        '//分页样例：[首页] [上页] [下页] [尾页] [页次：4/5页] [共86篇 20篇/页] 转到：_ 页
        '//标签：{N1} {N2} {N3} {N4} || 共:{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条
        If IsEmpty(TempStr) Then
            ErrInfo = ErrInfo&"模板为空..."
            StyleB = ErrInfo
            Exit Function
        End If
        Dim M
        If TotalPage>1 Then
            If CurrPage>1 Then
                M = "<a href='"&UrlStr&"Page=1'>"&"首页"&"</a>"
                TempStr = Replace(TempStr, "{N1}", M)
                M = "<a href='"&UrlStr&"Page="&CurrPage -1&"'>"&"上一页"&"</a>"
                TempStr = Replace(TempStr, "{N2}", M)
                If CurrPage<TotalPage Then
                    M = "<a href='"&UrlStr&"Page="&CurrPage+1&"'>"&"下一页"&"</a>"
                    TempStr = Replace(TempStr, "{N3}", M)
                    M = "<a href='"&UrlStr&"Page="&TotalPage&"'>"&"尾页"&"</a>"
                    TempStr = Replace(TempStr, "{N4}", M)
                Else
                    TempStr = Replace(TempStr, "{N3}", "下一页")
                    TempStr = Replace(TempStr, "{N4}", "尾页")
                End If
            Else
                TempStr = Replace(TempStr, "{N1}", "首页")
                TempStr = Replace(TempStr, "{N2}", "上一页")
                M = "<a href='"&UrlStr&"Page="&CurrPage+1&"'>"&"下一页"&"</a>"
                TempStr = Replace(TempStr, "{N3}", M)
                M = "<a href='"&UrlStr&"Page="&TotalPage&"'>"&"尾页"&"</a>"
                TempStr = Replace(TempStr, "{N4}", M)
            End If
        Else
            TempStr = Replace(TempStr, "{N1}", "首页")
            TempStr = Replace(TempStr, "{N2}", "上一页")
            TempStr = Replace(TempStr, "{N3}", "下一页")
            TempStr = Replace(TempStr, "{N4}", "尾页")
        End If
        T = TempStr
        T = Replace(T, "{N8}", TotalRecord)
        T = Replace(T, "{N6}", TotalPage)
        T = Replace(T, "{N5}", CurrPage)
        T = Replace(T, "{N7}", PageN)
        TempStr = T
        StyleA = TempStr
    End Function

    Private Function StyleB()
        '首页 |< 1 2 3 4 5 6 7 >| 尾页
        '//标签:{N1} {N2} {L}{N}{L/}{N3}{N4}
        If IsEmpty(TempStr) Then
            ErrInfo = ErrInfo&"模板为空..."
            StyleB = ErrInfo
            Exit Function
        End If
        Dim ForceNum, BackNum'//当前页的前面和后面显示个数
        ForceNum = 5
        BackNum = 4
        Dim M
        '//首页
        M = "<a href='"&UrlStr&"Page=1'>"&TempB(1)&"</a>"
        TempStr = Replace(TempStr, "{N1}", M)
        '//尾页
        M = "<a href='"&UrlStr&"Page="&TempB(6)&"'>"&TempB(4)&"</a>"
        TempStr = Replace(TempStr, "{N4}", M)
        '//前一页
        M = "|<"
        If CurrPage -1>= 1 Then
            M = "<a href='"&UrlStr&"Page="&CurrPage -1&"'>"&"|<"&"</a>"
        End If
        TempStr = Replace(TempStr, "{N2}", M)
        '//后一页
        M = ">|"
        If CurrPage+1<= TotalPage Then
            M = "<a href='"&UrlStr&"Page="&CurrPage+1&"'>"&">|"&"</a>"
        End If
        TempStr = Replace(TempStr, "{N3}", M)
        '//取出循环标签
        Dim N1, N2, N3, N4, N5, N6
        If InStr(TempStr, "{L}")>0 Then
            N1 = InStr(TempStr, "{L}")
        End If
        If InStr(TempStr, "{L/}")>0 Then
            N2 = InStr(TempStr, "{L/}")
        End If
        If N2<= N1 Then
            ErrInfo = ErrInfo&"循环标签出错..."
            StyleB = ErrInfo
            Exit Function
        End If
        N3 = Mid(TempStr, N1, N2 - N1 + 4)'//储存包括{L}{L/}循环标签的模板
        N4 = Replace(N3, "{L}", "")'//储存不包括{L}{L/}循环标签的模板
        N4 = Replace(N4, "{L/}", "")
        '//页码列表
        Dim FirstPageNum, LastPageNum
        If CurrPage - ForceNum<= 1 Then
            FirstPageNum = 1
            PageList = ""
        Else
            FirstPageNum = CurrPage - ForceNum
            PageList = "... ..."
        End If
        If CurrPage+BackNum>= TotalPage Then
            LastPageNum = TotalPage
            PageList_2 = ""
        Else
            LastPageNum = CurrPage+BackNum
            PageList_2 = "... ..."
        End If
        Dim I
        For I = FirstPageNum To LastPageNum
            If I = CurrPage * 1 Then
                N5 = Replace(N4, "{N}", "<b>"&I&"</b>")
                N6 = N6&N5
            Else
                M = "<a href='"&UrlStr&"Page="&I&"'>"&I&"</a>"
                N5 = Replace(N4, "{N}", M)
                N6 = N6&N5
            End If
        Next
        TempStr = Replace(TempStr, N3, N6)
        StyleB = TempStr
    End Function

    Private Function StyleC()
        '首页 |< |<< 1 2 3 4 5 6 7 >>| >| 尾页
        '//此风格在StyleB的基础上修改，增加两个标签：{N9}上10页 {N10}下10页
        '//标签:{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4}
        Dim T
        T = StyleB()
        '//前十页
        M = "|<<"
        If CurrPage -10>= 1 Then
            M = "<a href='"&UrlStr&"Page="&CurrPage -10&"'>"&"|<<"&"</a>"
        End If
        T = Replace(T, "{N9}", M)
        M = ">>|"
        If CurrPage+10<= TotalPage Then
            M = "<a href='"&UrlStr&"Page="&CurrPage+10&"'>"&">>|"&"</a>"
        End If
        T = Replace(T, "{N10}", M)
        StyleC = T
    End Function

    Private Function StyleD()
        '//此风格在StyleC的基础上修改
        '//首页 |< |<< 1 2 3 4 5 6 7 >>| >| 尾页　共{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条
        '//标签:{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4} 共{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条
        Dim T
        T = StyleC()
        T = Replace(T, "{N8}", TotalRecord)
        T = Replace(T, "{N6}", TotalPage)
        T = Replace(T, "{N5}", CurrPage)
        T = Replace(T, "{N7}", PageN)
        StyleD = T
    End Function

End Class
%>
