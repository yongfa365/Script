On Error Resume Next	'���������ִ����һ��

Function CheckUrl(strUrl)
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")
	objHTTP.Open "GET", strUrl, FALSE
	objHTTP.Send
	If objHTTP.statusText = "OK" Then
		CheckUrl = True
	Else
		CheckUrl = False
	End If
End Function

Set WSHShell = CreateObject("WScript.Shell") 
Favorites = WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites") 
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(Favorites)	'��fs��ȡ�����ļ����ϣ�strPathָ����ղؼ�Ŀ¼�µ��ļ���
Set fc = f.Files	'��f�ļ����ϴ���fc�У�fc�ͳ����ļ���������
For Each objFile in fc		'ѭ������fc�����е�ÿ���ļ�����
	If Lcase(Right(objFile.Name,4)) = ".url" Then	'ȷ���ҵ������ղؼ��е���ַ�����ļ�
		Set objStream = fs.OpenTextFile(strPath & objFile.Name)
			While Not objStream.AtEndOfStream	'��û�е����ļ�β����ʱ��ѭ��ִ��
				If "[InternetShortcut]"=objStream.ReadLine Then		'����1���ж��Ƿ�����ȷ��URL�ļ���ʽ
					targetPath = objStream.ReadLine		'����2�в�����targetPath����
					targetPath = Right(targetPath,len(targetPath)-4)	'���е������ǡ�URL=��ַ�������Ҫȥ��ǰ4���ַ��õ���ַ
					If CheckUrl(targetPath) = False Then	'����CheckUrl�������������ֵ��False�ͱ�ʾ����ַ�޷�����
						MsgBox objFile.Name & "�޷����ʣ������ѹ���" , vbinformation ,"��������"
					End If
				End If
			Wend
	End If
Next
