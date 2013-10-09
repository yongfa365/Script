'BrcIIS (Backup,Restore,Change Site IP) v1.0
'������ο���Դ��֮�ҵġ���ASP��̿���IIS����WEBվ�㡷��Adsutil.vbs�Ĵ���
'��Ҫ��лEnvymask��LuoLuo�İ�����лл
'С����ADSIҲ���Ǻ���Ϥ�����Դ�������ʲô����ģ�����������ָ�̣�лл
'Codz by BlackFox,My QQ:6858849 E-mail:ym2236@163.com WebSite:fox.he100.com
'Base On Adsutil.vbs

Option Explicit
'On Error Resume Next

Dim objShell,objArg,objInStream,objOutStream,intSiteIndex,objW3svc,ObjChildObject,strChildObjectName
Dim objIIs,objIIsWeb,objFso,objBackupFile,arrServerBindings,strServerComment,strMaxConnections
Dim strPath,strBackupFile,strFileLine,arrSiteInfo,intCreateWebStatus,strOldIP,strNewIP,intSiteCount
Dim intBindsIndex,intBindsArrID,intChoice,strSearchMode,strSearchedSite,strSaveToBackup,strResult
Dim strSaveFileName,strOverWrite,objResultFile,strIpString,intSearched,strDomainString,strPathString
Dim strDeleteMode,intDeleteId,arrSiteId(999),strCommentString,intCountDeleteSite

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

If (LCase(Right(Wscript.FullName , 11))="wscript.exe") Then
   Set objShell=Wscript.CreateObject("Wscript.Shell")
   objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&Wscript.ScriptFullName&chr(34))
   Wscript.Quit
End If

Set objOutStream = Wscript.StdOut
Set objInStream  = Wscript.StdIn

Set objArg = Wscript.Arguments
If objArg.Count <= 0 Then
HELP()
Wscript.Quit
End If

Select Case UCase(objArg(0))
Case "BACKUP"
HELP()
BackupSite()
Case "RESTORE"
HELP()
RestoreSite()
Case "CHANGEIP"
HELP()
ChangeSiteIP()
Case "CHANGE_APP"
HELP()
Change_AppIsolated()
Case "SEARCH"
HELP()
Search()
Case "DELETESITE"
HELP()
DeleteSite()
End Select

Function BackupSite() '����վ��ģ��
	objOutStream.Write "��������������վ����Ϣ���ļ���:"
	strBackupFile = objInStream.ReadLine
	intSiteIndex = 1 'վ��ID������
	Wscript.Echo vbCrLf & "��ʼ����վ����Ϣ����" & vbCrLf
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '����IIS����
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�ObjChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name)
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			strServerComment  = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
			strMaxConnections = objIIs.MaxConnections '��IIS����վ��������������ֵ������MaxConnections
			Set objIIsWeb = objIIs.GetObject("IIsWebVirtualDir","Root")
			strPath = objIIsWeb.Path
			Set objFso = CreateObject("Scripting.FileSystemObject")
			If (objFso.fileexists(strBackupFile)) Then
				Set objBackupFile = objFso.OpenTextFile(strBackupFile,ForAppending,True)
				objBackupFile.WriteLine(strServerComment  &vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath)
				objBackupFile.Close
				Wscript.Echo "���ڱ���վ�� " & strServerComment & "!"
			Else
				Set objBackupFile = objFso.CreateTextFile(strBackupFile,True)
				objBackupFile.WriteLine(strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath)
				objBackupFile.Close
				Wscript.Echo "���ڱ���վ�� " & strServerComment & "!"
			End If
		intSiteIndex = intSiteIndex + 1
		End If
	Next
	Wscript.Echo vbCrLf & "һ�������� " & intSiteIndex - 1 & "��վ����Ϣ!" & vbCrLf
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Set objFso         = Nothing
	Set objBackupFile  = Nothing
End Function

Function RestoreSite() '�ָ�վ��ģ��
	objOutStream.Write "�����������ָ�վ����Ϣ���ļ���:"
	strBackupFile = objInStream.ReadLine
	Wscript.Echo vbCrLf & "��ʼ�ָ�վ����Ϣ����" & vbCrLf
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If Not (objFso.fileexists(strBackupFile)) Then
	Wscrip.Echo "û���ҵ��ļ� "&strBackupFile&"!"&vbCrLf&"��ȷ���ļ���λ��!"
	Wscript.Quit
	End If
	Set objBackupFile = objFso.OpenTextFile(strBackupFile,1)
	intSiteCount = 1
	Do While objBackupFile.AtEndOfStream <> True
		strFileLine = objBackupFile.ReadLine
		arrSiteInfo = split(strFileLine,vbTab)
		intCreateWebStatus = CreateWebServer(arrSiteInfo(0),split(arrSiteInfo(1),","),arrSiteInfo(2),arrSiteInfo(3))
		If intCreateWebStatus = 1 Then
			Wscript.Echo "����վ�� " & arrSiteInfo(0) & " �ɹ�"
			intSiteCount = intSiteCount + 1
		Else
			Wscript.Echo "����վ�� " & arrSiteInfo(0) & " ʧ��"
		End If
	Loop
	objBackupFile.Close
	Wscript.Echo vbCrLf & "һ���ָ��� " & intSiteCount - 1 & " ��վ����Ϣ!" & vbCrLf
	Set objBackupFile = Nothing
	Set objFso        = Nothing
End Function


Function ChangeSiteIP() '�޸�վ��IPģ��
	'*********************************************
	'��ȡ����IP��һ����ԭ���ģ��Ǹ����µ�
	'*********************************************
	objOutStream.Write "������վ��ԭIP��"
	strOldIP = objInStream.ReadLine
	objOutStream.Write "������վ����IP��"
	strNewIP = objInStream.ReadLine
	'*********************************************
	Wscript.Echo vbCrLf & "��ʼ�޸ġ���" & vbCrLf
	Set objW3svc = GetObject("IIS://localhost/w3svc") '����IIS����
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�objChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			strServerComment  = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
			Wscript.Echo "�����޸�վ��  " & strServerComment & "  ��IP" '�������д�ӡ�����޸�IP��վ������
			intBindsIndex = UBound(arrServerBindings) + 1 '����վ����԰󶨶���������˿ں�IP��������Ҫ�ж�һ�����˶���IP
			For intBindsArrID = 1 To intBindsIndex
				If instr(arrServerBindings(intBindsArrID - 1) , strOldIP) Then
					arrServerBindings(intBindsArrID - 1) = Replace(arrServerBindings(intBindsArrID - 1) , strOldIP , strNewIP) '�������ڰ���oldip���ַ����滻��newip���ַ���
				End If
			Next
			objIIs.ServerBindings = arrServerBindings '��iis�����ServerBindings�����޸�Ϊ�滻��IP������
			objIIs.setinfo 'ʹ�滻����������Ч
		End If
	Next
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Wscript.Echo vbCrLf & "�޸����!"
End Function

Function Change_AppIsolated() '�޸�Ӧ�ó��򱣻�ģ��
	objOutStream.Write "ѡ������Ҫ���õ�(0,��_IIS���� 2,��_���õ� 3,��_������,Ĭ������Ϊ2):"
	intChoice = objInStream.ReadLine
	If Not IsNumeric(intChoice) Then
		Wscript.Echo "����������Ĳ�������!"
		Wscript.Quit
	End If
	intSiteIndex = 1 'վ��ID������
	Wscript.Echo vbCrLf & "��ʼ�޸ġ���" & vbCrLf
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '����IIS����
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�ȡ����objChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer" , objChildObject.Name)
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			strServerComment = objIIs.ServerComment
			set objIIsWeb = objIIs.GetObject("IIsWebVirtualDir","Root")
			objIIsWeb.AppIsolated = int(intChoice)
			objIIsWeb.SetInfo
			Wscript.Echo "����վ�� " & strServerComment & " ���!"
		End If
	Next
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Wscript.Echo vbCrLf & "�������!"
End Function

Function Search() '����ģ��
	Wscript.Echo "��ѡ��������ʽ��" & vbCrLf
	Wscript.Echo "1.����IP��ѯվ����Ϣ"
	Wscript.Echo "2.����������ѯվ����Ϣ"
	Wscript.Echo "3.����վ�����·����ѯվ����Ϣ" & vbCrLf
	objOutStream.Write "��ѡ������Ҫ�Ĳ�ѯ��ʽ��1 2 3����"
	strSearchMode = objInStream.ReadLine

	Select Case strSearchMode
	Case "1"
	Use_IP_Search()
	Case "2"
	Use_Domain_Search()
	Case "3"
	Use_Path_Search()
	Case Else
	Wscript.Echo String(30," ") & "�������,����������" & vbCrLf
	Search()
	End Select
End Function

Function Use_IP_Search() '����IP����վ����Ϣ
	objOutStream.Write "��������Ҫ������IP(֧��ģ������)��"
	strIpString = objInStream.ReadLine
	Wscript.Echo "��ʼ��������" & vbCrLf
	intSiteIndex = 1
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '����IIS����
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�objChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '����վ����԰󶨶���������˿ں�IP��������Ҫ�ж�һ�����˶���IP
			intSearched=0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strIpString) Then
					intSearched=1
					Exit For
				End If
			Next
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				strMaxConnections = objIIs.MaxConnections
				Wscript.Echo "վ��������" & strServerComment
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '��ӡ����������վ����Ϣ
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "δ���ҵ�վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "�����ҵ�" & intSiteIndex - 1 & "��վ��!"
		Wscript.Echo "Ĭ��ֻ��ʾ��վ���������������Ҫ��ϸ��Ϣ��������������Ϊ�ļ�!"
		objOutStream.Write "�Ƿ�վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function

Function Use_Domain_Search() '��������ģ��
	objOutStream.Write "��������Ҫ����������֧��ģ��������"
	strDomainString = objInStream.ReadLine
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	Wscript.Echo "��ʼ��������" & vbCrLf
	intSiteIndex = 1
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�objChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '����վ����԰󶨶���������˿ں�IP��������Ҫ�ж�һ�����˶���IP
			intSearched=0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strDomainString) Then
					intSearched=1
					Exit For
				End If
			Next
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				strMaxConnections = objIIs.MaxConnections
				Wscript.Echo "վ��������" & strServerComment
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '��ӡ����������վ����Ϣ
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "δ���ҵ�վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "�����ҵ�" & intSiteIndex - 1 & "��վ��!"
		Wscript.Echo "Ĭ��ֻ��ʾ��վ���������������Ҫ��ϸ��Ϣ��������������Ϊ�ļ�!"
		objOutStream.Write "�Ƿ�վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing

End Function

Function Use_Path_Search() '��������ģ��
	objOutStream.Write "��������Ҫ�����ľ���·��(֧��ģ������)��"
	strPathString = objInStream.ReadLine
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	Wscript.Echo "��ʼ��������" & vbCrLf
	intSiteIndex = 1
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '�ж�objChildObject.Name�ǲ�������
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If instr(objIIsWeb.Path , strPathString) Then
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
				strMaxConnections = objIIs.MaxConnections
				strPath=objIIsWeb.Path
				Wscript.Echo "վ��������" & strServerComment
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '��ӡ����������վ����Ϣ
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "δ���ҵ�վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "�����ҵ�" & intSiteIndex - 1 & "��վ��!"
		Wscript.Echo "Ĭ��ֻ��ʾ��վ���������������Ҫ��ϸ��Ϣ��������������Ϊ�ļ�!"
		objOutStream.Write "�Ƿ�վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function

Function DeleteSite()
	Wscript.Echo vbCrLf & "��ѡ������ɾ���ķ�ʽ:" & vbCrLf
	Wscript.Echo "1.����IPɾ��վ��" & vbCrLf
	Wscript.Echo "2.��������ɾ��վ��" & vbCrlf
	Wscript.Echo "3.����վ������ɾ��վ��" & vbCrLf
	Wscript.Echo "4.ɾ������վ����Ϣ" & vbCrLf
	objOutStream.Write "��ѡ������Ҫ��ɾ����ʽ��1 2 3 4����"
	strDeleteMode = objInStream.ReadLine

	Select Case strDeleteMode
		Case "1"
		Delete_By_Ip()
		Case "2"
		Delete_By_Domain()
		Case "3"
		Delete_By_SiteComment()
		Case "4"
		Delete_All_Site()
		Case Else
		Wscript.Echo vbCrLf & String(30," ") & "�������,����������!"
		DeleteSite()
	End Select
End Function

Function Delete_By_Ip()
	objOutStream.Write "��������Ҫɾ����վ���IP(֧��ģ����ѯ):"
	strIpString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "��ʼ��������" & vbCrLf
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	intSiteIndex = 0
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then
			arrSiteId(intSiteIndex) = CLng(objChildObject.Name)
			intSiteIndex = intSiteIndex + 1
		End If
	Next
	intCountDeleteSite = 0
	For intDeleteId = 0 To intSiteIndex - 1
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '����վ����԰󶨶���������˿ں�IP��������Ҫ�ж�һ�����˶���IP
			intSearched = 0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strIpString) Then
					intSearched=1
					Exit For
				End If
			Next
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "ɾ��վ��: " & strServerComment & "���!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "δ���ҵ�Ҫɾ����վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "��ɾ��" & intCountDeleteSite & "��վ��!"
		objOutStream.Write "�Ƿ���ɾ����վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function

Function Delete_By_Domain()
	objOutStream.Write "��������Ҫɾ����վ�������(֧��ģ����ѯ):"
	strDomainString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "��ʼ��������" & vbCrLf
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	intSiteIndex = 0
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then
			arrSiteId(intSiteIndex) = CLng(objChildObject.Name)
			intSiteIndex = intSiteIndex + 1
		End If
	Next
	intCountDeleteSite = 0
	For intDeleteId = 0 To intSiteIndex - 1
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '����վ����԰󶨶���������˿ں�IP��������Ҫ�ж�һ�����˶���IP
			intSearched = 0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strDomainString) Then
					intSearched=1
					Exit For
				End If
			Next
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "ɾ��վ��: " & strServerComment & "���!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "δ���ҵ�Ҫɾ����վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "��ɾ��" & intCountDeleteSite & "��վ��!"
		objOutStream.Write "�Ƿ���ɾ����վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function

Function Delete_By_SiteComment
	objOutStream.Write "��������Ҫɾ����վ�������(֧��ģ����ѯ):"
	strCommentString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "��ʼ��������" & vbCrLf
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	intSiteIndex = 0
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then
			arrSiteId(intSiteIndex) = CLng(objChildObject.Name)
			intSiteIndex = intSiteIndex + 1
		End If
	Next
	intCountDeleteSite = 0
	For intDeleteId = 0 To intSiteIndex - 1
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			If Instr(Cstr(objIIs.ServerComment) , strCommentString) or Cstr(objIIs.ServerComment) = strCommentString  Then
				intSearched=1
			End If
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment
				arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "ɾ��վ��: " & strServerComment & "���!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "δ���ҵ�Ҫɾ����վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "��ɾ��" & intCountDeleteSite & "��վ��!"
		objOutStream.Write "�Ƿ���ɾ����վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function

Function Delete_All_Site()
	Wscript.Echo vbCrLf & "��ʼɾ������" & vbCrLf
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	intSiteIndex = 0
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then
			arrSiteId(intSiteIndex) = CLng(objChildObject.Name)
			intSiteIndex = intSiteIndex + 1
		End If
	Next
	intCountDeleteSite = 0
	For intDeleteId = 0 To intSiteIndex - 1
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '����IIS����վ�����
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '��IIS����վ��󶨵�IP���˿ڡ�������ֵ��������ServerBindings
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strServerComment = objIIs.ServerComment '��IIS����վ������Ƹ�ֵ������ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "ɾ��վ��: " & strServerComment & "���!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intCountDeleteSite = intCountDeleteSite + 1
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "δ���ҵ�Ҫɾ����վ����Ϣ!"
	Else
		Wscript.Echo vbCrLf & "��ɾ��" & intCountDeleteSite & "��վ��!"
		objOutStream.Write "�Ƿ���ɾ����վ����Ϣ��Ϊ�����ļ�����YES��NO����"
		strSaveToBackup=objInStream.ReadLine
		If UCase(strSaveToBackup)="YES" Then
			SaveToFile(strSearchedSite)
		End If
	End If
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
End Function


Function SaveToFile(strResult) '�����ļ�
	objOutStream.Write "�����뱣����ļ�����"
	strSaveFileName=objInStream.ReadLine
	Set objFso=CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(strSaveFileName) Then
		Wscript.Echo "�ļ�" & strSaveFileName & "�Ѿ�����!"
		objOutStream.Write "�Ƿ�׷��д���ļ���(Yes,No)"
		strOverWrite = objInStream.ReadLine
		If UCase(strOverWrite) = "YES" Then
			Set objResultFile=objFso.OpenTextFile(strSaveFileName , ForAppending , True)
			objResultFile.Write strResult
			objResultFile.Close
			Wscript.Echo "д���ļ�" & strSaveFileName & "���!"
		End If

		If UCase(strOverWrite) = "NO" Then
			objOutStream.Write "���������ļ���"
			strSaveFileName = objInStream.ReadLine
			Set objResultFile = objFso.CreateTextFile(strSaveFileName,True)
			objResultFile.Write strResult
			objResultFile.Close
			Wscript.Echo "�����ļ� " & strSaveFileNmae & " ���!"
		End If
	Else
		Set objResultFile = objFso.CreateTextFile(strSaveFileName , True)
		objResultFile.Write strResult
		objResultFile.Close
		Wscript.Echo "�����ļ� " & strSaveFileName & " ���!"
	End If
	Set objFso = Nothing
End Function


Function CreateWebServer(strServerComment,arrServerBindings,strMaxConnections,strPath)'����վ��
	On Error Resume Next
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	intSiteIndex = 999

	Do While IsObject(objW3svc.GetObject("IIsWebServer",intSiteIndex))
		If Err.Number <> 0 Then
			'Wscript.Echo Err.Description
			Err.Clear()
			Exit Do
		End If
		intSiteIndex = intSiteIndex - 1
	Loop

	Set objIIs = objW3svc.Create("IIsWebServer",intSiteIndex)
	If Err.Number <> 0 Then
		Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
		Err.Clear()
		CreateWebServer = 0
		Exit Function
	End If

	objIIs.ServerSize = 1
	objIIs.ServerComment    = strServerComment
	objIIs.ServerBindings   = arrServerBindings
	objIIs.MaxConnections   = strMaxConnections
	objIIs.EnableDefaultDoc = True
	objIIs.SetInfo
	Set objIIsWeb = objIIs.Create("IIsWebVirtualDir", "Root")
	If Err.Number <> 0 Then
		Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
		Err.Clear()
		CreateWebServer = 0
		Exit Function
	End If
	objIIsWeb.Path              = strPath
	objIIsWeb.AccessRead        = True
	objIIsWeb.AccessWrite       = False
	objIIsWeb.EnableDirBrowsing = False 
	objIIsWeb.EnableDefaultDoc  = True 
	objIIsWeb.AccessScript      = True
	objIIsWeb.AppIsolated       = 2
	objIIsWeb.AppCreate2  		  2
	objIIsWeb.AppFriendlyName   = "Ĭ��Ӧ�ó���" 
	objIIsWeb.SetInfo
	Set objW3svc  = Nothing
	Set objIIs    = Nothing
	Set objIIsWeb = Nothing
	CreateWebServer = 1
End Function

Function HELP()
	Wscript.Echo vbCrLf & "                          IIS�ܹ� V1.6 By BlackFox"
	Wscript.Echo "                            http://fox.he100.com" & vbCrLf
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Backup ������IISվ����Ϣ��һ���ı��ļ�����"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Restroe ����һ���ı��ļ��ָ�IISվ����Ϣ����"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs ChangeIP ���޸�����IISվ��IP����"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Change_App ���޸�����IISվ��Ӧ�ó��򼶱𣡣�"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Search ��������������վ����Ϣ����"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs DeleteSite (������������ɾ��վ��)" & vbCrLf
End Function

'�뵽���¹���
'���ݹ�������һ������վ�����ݵĹ��ܣ�ͨ��RAR�Զ������Ĭ�ϴ����վ����ϼ�Ŀ¼����ѡ������ָ����Ŀ¼
'���ӽ�����������Ŀ¼�Ĺ���
'����ɾ���󶨵�ָ��IP
'��������ʱ�ж�IIS�Ƿ�����
'�ָ�վ�����·����ACL��Ϣ