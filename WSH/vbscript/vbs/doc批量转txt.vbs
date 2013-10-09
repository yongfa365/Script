Sub DocConvert(strFile)
	wdFormatText = 2	'要传递给SaveAs的参数，表示文本类型
	Dim wApp
	Set wApp = CreateObject("Word.Application")	'生成Word应用程序对象
	wApp.Visible = False	'调用对象过程中不显示Word界面
	Set myDoc = wApp.Documents.open(strFile)	'打开过程参数指向的文件
	myDoc.SaveAs strFile & ".txt" ,wdFormatText	'将DOC文件保存为TXT，文件名用原DOC文件名加上“.txt”后缀
	wApp.Quit	'退出Word程序
End Sub

Sub ProcessFiles(strPath)
	Dim fs,f,fc,objFile
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set f = fs.GetFolder(strPath)	'把fs获取到的文件集合（strPath指向的目录下的文件）
	Set fc = f.Files	'将f文件集合存入fc中，fc就成了文件对象数组
	For Each objFile in fc		'循环处理fc数组中的每个文件对象
		If Lcase(Right(objFile.Name,4)) = ".doc" Then	'只有DOC文件才做转换，取文件名右侧4位字母判断（判断前转成小写以确保匹配）
			DocConvert strPath & objFile.Name	'调用DocConvert过程，参数是文件的全路径（路径＋文件对象的文件名）
		End If
	Next
End Sub


Set objArgs = WScript.Arguments		'通过WScript对象的Arguments获取参数对象数组
If Right(objArgs(0),1) <> "\" Then	'只使用第1个（索引是0）参数，并保证参数中的路径最后包含“\”
	ProcessFiles objArgs(0) & "\"
Else
	ProcessFiles objArgs(0)
End If