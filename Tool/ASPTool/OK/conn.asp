<!--#include file="dbpath.asp" -->
<%
'On Error Resume Next
connlink = "ACCESS"
Select Case connlink
    Case "ACCESS"
        
        '*********************************************************************
        dbpath = server.mappath(dbpath)
        
        connstr = "provider=microsoft.jet.oledb.4.0;data source=" & dbpath
        Set conn = server.CreateObject("adodb.connection")
        conn.Open connstr
        '*********************************************************************
    Case "MSSQL"
        '*********************************************************************
        SqlDbHost = "." '修改实际的数据库服务器地址
        SqlDbName = "test001" '修改为实际的数据库名
        SqlUserName = "sa" '修改为实际的数据库用户名
        SqlUserPass = "jhsb_fxzx_jsz" '修改为实际的数据库密码
        '*********************************************************************
        
        Set conn = Server.CreateObject("ADODB.Connection")
        'SQL链接字符串
        connstr = "Provider=SQLOLEDB.1; Persist Security Info=True; Data Source=" & SqlDbHost & "; Initial Catalog="&SqlDbName&"; User ID="&SqlUserName&"; Password=" & SqlUserPass
        conn.Open connstr
End Select

'Set rs = Server.CreateObject("ADODB.Recordset")


Function SafeRequest(ParaValue)
    ParaValue = Trim(Request(ParaValue))
    If ParaValue = "" Then
        SafeRequest = ""
        Exit Function
    End If
    '要过滤的字符以","隔开
    LockValue = "',Select,Update,Delete,insert,Count(,drop table,truncate,Asc(,Mid(,char(,xp_cmdshell,exec master,net localgroup administrators, And ,net user, Or "
    LockValue = Split(LockValue, ",")
    '判断是否有注入
    For i = 0 To UBound(LockValue)
        If InStr(LCase(ParaValue), LockValue(i))>0 Then
            errmsg = 1
            Exit For
        End If
    Next
    '注入处理
    If errmsg = 1 Then
        Response.Write "<script language='javascript'>alert('可疑的SQL注入请求!');window.history.go(-1);</script>"
    Else
        SafeRequest = ParaValue
    End If
End Function


Function isnum(Str)
    str1 = Trim(request(Str))
    If str1 = "" Or IsNull(str1) Then
        isnum = 0
    Else
        isnum = str1
    End If
End Function



Sub CloseConn()
    conn.Close
    Set conn = Nothing
End Sub

'禁止外部提交

Function safein(Url)
    server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
    server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
    If Mid(server_v1, 8, Len(server_v2))<>server_v2 Then
        response.Write "<script>alert('警告！你正在从外部提交数据！！请立即终止！！\n你的IP已经被记录，如果再次来');</script>"
        Response.Redirect Url
    End If
End Function


%>
