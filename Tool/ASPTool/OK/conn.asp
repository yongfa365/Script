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
        SqlDbHost = "." '�޸�ʵ�ʵ����ݿ��������ַ
        SqlDbName = "test001" '�޸�Ϊʵ�ʵ����ݿ���
        SqlUserName = "sa" '�޸�Ϊʵ�ʵ����ݿ��û���
        SqlUserPass = "jhsb_fxzx_jsz" '�޸�Ϊʵ�ʵ����ݿ�����
        '*********************************************************************
        
        Set conn = Server.CreateObject("ADODB.Connection")
        'SQL�����ַ���
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
    'Ҫ���˵��ַ���","����
    LockValue = "',Select,Update,Delete,insert,Count(,drop table,truncate,Asc(,Mid(,char(,xp_cmdshell,exec master,net localgroup administrators, And ,net user, Or "
    LockValue = Split(LockValue, ",")
    '�ж��Ƿ���ע��
    For i = 0 To UBound(LockValue)
        If InStr(LCase(ParaValue), LockValue(i))>0 Then
            errmsg = 1
            Exit For
        End If
    Next
    'ע�봦��
    If errmsg = 1 Then
        Response.Write "<script language='javascript'>alert('���ɵ�SQLע������!');window.history.go(-1);</script>"
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

'��ֹ�ⲿ�ύ

Function safein(Url)
    server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
    server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
    If Mid(server_v1, 8, Len(server_v2))<>server_v2 Then
        response.Write "<script>alert('���棡�����ڴ��ⲿ�ύ���ݣ�����������ֹ����\n���IP�Ѿ�����¼������ٴ���');</script>"
        Response.Redirect Url
    End If
End Function


%>
