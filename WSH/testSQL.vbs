'***********************SQL Server**************************************
SqlDbHost = "database.cbdnic.com" '�޸�ʵ�ʵ����ݿ��������ַ
SqlDbName = "hangye" '�޸�Ϊʵ�ʵ����ݿ���
SqlUserName = "hangye" '�޸�Ϊʵ�ʵ����ݿ��û���
SqlUserPass = "hangyehangye" '�޸�Ϊʵ�ʵ����ݿ�����
'*********************************************************************
Set conn = CreateObject("ADODB.Connection")
'SQL�����ַ���
connstr = "Provider=SQLOLEDB.1; Persist Security Info=True; Data Source=" & SqlDbHost & "; Initial Catalog="&SqlDbName&"; User ID="&SqlUserName&"; Password=" & SqlUserPass
conn.Open connstr
If Err>0 Then
    MsgBox("error")
Else
    MsgBox("success")
End If
