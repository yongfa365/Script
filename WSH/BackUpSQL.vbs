Dim OKstr

Set conn = CreateObject("ADODB.Connection")
conn.Open " Provider=SQLOLEDB;data source=127.0.0.1;initial catalog=master;user id=sa;password=;"
sql = "select  db_name(dbid) dbname,name,filename from master..sysaltfiles"
'sql = "select  db_name(dbid) dbname,name,filename from master..sysaltfiles where db_name(dbid) not in('master','tempdb','model','msdb')"
Set rs = CreateObject("adodb.recordset")
rs.Open sql, conn, 1,1

Do While Not rs.EOF

    i = 1
    OKstr = OKstr & "EXEC sp_attach_db @dbname = '" & rs("dbname") & "'" & vbCrLf
    OKstr = OKstr & ",@filename"&i&" = '" & rs("filename") & "'" & vbCrLf
    
    db_name = rs("dbname")
    
    Do While Not rs.EOF
        rs.movenext
        if rs.eof then exit do
        If rs("dbname") = db_name Then
            i = i + 1
            OKstr = OKstr & ",@filename"&i&" = '"&rs("filename")&"'" & vbCrLf
        Else
            Exit Do
        End If
    Loop
    OKstr = OKstr & vbCrLf
Loop


CreateFile "OK.txt", OKstr

Function CreateFile(FileName, Content)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fd = FSO.CreateTextFile(FileName, True)
    fd.WriteLine Content
End Function











'set conn=createobject("adodb.connection")
'conn.open "Data Source=.;Initial Catalog=master;Integrated Security=True"
'sql="select  db_name(dbid) dbname,name,filename from master..sysaltfiles"
'Set rs= CreateObject("ADODB.Recordset")
'rs.Open sql,conn,1,3
'for i=1 to rs.recordcount
'
'
'next

'select  db_name(dbid) dbname,name,filename from master..sysaltfiles where db_name(dbid) not in('master','tempdb','model','msdb')
