'/*=========================================================================
' * Intro       
' * FileName    ACCESS���ݿ�ѹ��.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/item/ACCESS-CompactDatabase-97-2000-2003-yongfa365.html
' * LastModify  2007-09-25 13:27:17
' *==========================================================================*/

'====================ѹ�����ݿ� =========================
Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count - 1

	dbpath = objArgs(I)
	boolIs97 = msgbox("97�����ݿ�?",vbYesNo)

	If dbpath <> "" Then
	CompactDB dbpath,boolIs97
	End If

Next
'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
	fso.CopyFile dbpath,strDBPath & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
	End If

fso.CopyFile strDBPath & "temp1.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
fso.DeleteFile(strDBPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing
	CompactDB = "�ɹ�ѹ��! " & vbCrLf
Else
	CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
End If
Msgbox CompactDB 
End Function
