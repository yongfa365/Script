<%
connlink="1"
select case connlink
case "1"
'*********************************************************************
dbpath=server.mappath("/200712345yongfa365/#%20#%20yongfa.mdb")  
connstr= "provider=microsoft.jet.oledb.4.0;data source=" & dbpath 
set conn=server.createobject("adodb.connection")  
conn.open connstr
'*********************************************************************
case "2"
'*********************************************************************
SqlDbHost="127.0.0.1"  '�޸�ʵ�ʵ����ݿ��������ַ
SqlDbName="datebase"  '�޸�Ϊʵ�ʵ����ݿ���
SqlUserName="username"  '�޸�Ϊʵ�ʵ����ݿ��û���
SqlUserPass="password" '�޸�Ϊʵ�ʵ����ݿ�����
'*********************************************************************

Set conn = Server.CreateObject("ADODB.Connection")
'SQL�����ַ���
connstr="Provider=SQLOLEDB.1; Persist Security Info=True; Data Source=" & SqlDbHost & "; Initial Catalog="&SqlDbName&"; User ID="&SqlUserName&"; Password=" & SqlUserPass
conn.Open connstr

end select

'Set rs = Server.CreateObject("ADODB.Recordset")


Function SafeRequest(ParaValue) 
ParaValue=Trim(Request(ParaValue))
If ParaValue = "" Then
	SafeRequest = ""
	Exit Function
End If
'Ҫ���˵��ַ���","����
LockValue="',Select,Update,Delete,insert,Count(,drop table,truncate,Asc(,Mid(,char(,xp_cmdshell,exec master,net localgroup administrators, And ,net user, Or "
LockValue=Split(LockValue,",")
'�ж��Ƿ���ע��
For i=0 To UBound(LockValue)
If InStr(LCase(ParaValue),LockValue(i))>0 Then
errmsg=1
Exit For
End If
Next
'ע�봦��
If errmsg=1 Then
	errmsg= "<script language='javascript'>alert('���ɵ�SQLע������!');"
	errmsg=errmsg & "window.history.go(-1);" 
	errmsg=errmsg & "</script>"
	Response.Write errmsg
Else
	SafeRequest=ParaValue
End If
End Function


function isnum(str)
str1=trim(request(str))
if str1="" or ISNULL(str1) then
isnum=0
else
isnum=str1
end if
end function



sub CloseConn()
	conn.close
	set conn=nothing
end sub

'��ֹ�ⲿ�ύ
function safein(Url)
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if  mid(server_v1,8,len(server_v2))<>server_v2  then
response.write "<script>alert('���棡�����ڴ��ⲿ�ύ���ݣ�����������ֹ����\n���IP�Ѿ�����¼������ٴ���');</script>"
Response.Redirect Url
end if
end function

%>

<!--#include file="function.asp"-->