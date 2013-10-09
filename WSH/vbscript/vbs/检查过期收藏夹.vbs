On Error Resume Next	'发生错误就执行下一行

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
Set f = fs.GetFolder(Favorites)	'把fs获取到的文件集合（strPath指向的收藏夹目录下的文件）
Set fc = f.Files	'将f文件集合存入fc中，fc就成了文件对象数组
For Each objFile in fc		'循环处理fc数组中的每个文件对象
	If Lcase(Right(objFile.Name,4)) = ".url" Then	'确保找到的是收藏夹中的网址链接文件
		Set objStream = fs.OpenTextFile(strPath & objFile.Name)
			While Not objStream.AtEndOfStream	'在没有到达文件尾部的时候都循环执行
				If "[InternetShortcut]"=objStream.ReadLine Then		'读第1行判断是否是正确的URL文件格式
					targetPath = objStream.ReadLine		'读第2行并存入targetPath变量
					targetPath = Right(targetPath,len(targetPath)-4)	'这行的内容是“URL=网址”，因此要去除前4个字符得到网址
					If CheckUrl(targetPath) = False Then	'调用CheckUrl函数，如果返回值是False就表示此网址无法访问
						MsgBox objFile.Name & "无法访问，可能已过期" , vbinformation ,"过期链接"
					End If
				End If
			Wend
	End If
Next
