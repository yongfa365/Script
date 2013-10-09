Sub DocConvert(strFile)
	wdFormatText = 2	'Ҫ���ݸ�SaveAs�Ĳ�������ʾ�ı�����
	Dim wApp
	Set wApp = CreateObject("Word.Application")	'����WordӦ�ó������
	wApp.Visible = False	'���ö�������в���ʾWord����
	Set myDoc = wApp.Documents.open(strFile)	'�򿪹��̲���ָ����ļ�
	myDoc.SaveAs strFile & ".txt" ,wdFormatText	'��DOC�ļ�����ΪTXT���ļ�����ԭDOC�ļ������ϡ�.txt����׺
	wApp.Quit	'�˳�Word����
End Sub

Sub ProcessFiles(strPath)
	Dim fs,f,fc,objFile
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set f = fs.GetFolder(strPath)	'��fs��ȡ�����ļ����ϣ�strPathָ���Ŀ¼�µ��ļ���
	Set fc = f.Files	'��f�ļ����ϴ���fc�У�fc�ͳ����ļ���������
	For Each objFile in fc		'ѭ������fc�����е�ÿ���ļ�����
		If Lcase(Right(objFile.Name,4)) = ".doc" Then	'ֻ��DOC�ļ�����ת����ȡ�ļ����Ҳ�4λ��ĸ�жϣ��ж�ǰת��Сд��ȷ��ƥ�䣩
			DocConvert strPath & objFile.Name	'����DocConvert���̣��������ļ���ȫ·����·�����ļ�������ļ�����
		End If
	Next
End Sub


Set objArgs = WScript.Arguments		'ͨ��WScript�����Arguments��ȡ������������
If Right(objArgs(0),1) <> "\" Then	'ֻʹ�õ�1����������0������������֤�����е�·����������\��
	ProcessFiles objArgs(0) & "\"
Else
	ProcessFiles objArgs(0)
End If