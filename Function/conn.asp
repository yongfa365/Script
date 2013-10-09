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
SqlDbHost="127.0.0.1"  '修改实际的数据库服务器地址
SqlDbName="datebase"  '修改为实际的数据库名
SqlUserName="username"  '修改为实际的数据库用户名
SqlUserPass="password" '修改为实际的数据库密码
'*********************************************************************

Set conn = Server.CreateObject("ADODB.Connection")
'SQL链接字符串
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
'要过滤的字符以","隔开
LockValue="',Select,Update,Delete,insert,Count(,drop table,truncate,Asc(,Mid(,char(,xp_cmdshell,exec master,net localgroup administrators, And ,net user, Or "
LockValue=Split(LockValue,",")
'判断是否有注入
For i=0 To UBound(LockValue)
If InStr(LCase(ParaValue),LockValue(i))>0 Then
errmsg=1
Exit For
End If
Next
'注入处理
If errmsg=1 Then
	errmsg= "<script language='javascript'>alert('可疑的SQL注入请求!');"
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

'禁止外部提交
function safein(Url)
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if  mid(server_v1,8,len(server_v2))<>server_v2  then
response.write "<script>alert('警告！你正在从外部提交数据！！请立即终止！！\n你的IP已经被记录，如果再次来');</script>"
Response.Redirect Url
end if
end function

%>

<!--#include file="function.asp"-->