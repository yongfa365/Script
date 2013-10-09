'***********************SQL Server**************************************
SqlDbHost = "database.cbdnic.com" '修改实际的数据库服务器地址
SqlDbName = "hangye" '修改为实际的数据库名
SqlUserName = "hangye" '修改为实际的数据库用户名
SqlUserPass = "hangyehangye" '修改为实际的数据库密码
'*********************************************************************
Set conn = CreateObject("ADODB.Connection")
'SQL链接字符串
connstr = "Provider=SQLOLEDB.1; Persist Security Info=True; Data Source=" & SqlDbHost & "; Initial Catalog="&SqlDbName&"; User ID="&SqlUserName&"; Password=" & SqlUserPass
conn.Open connstr
If Err>0 Then
    MsgBox("error")
Else
    MsgBox("success")
End If
