'BrcIIS (Backup,Restore,Change Site IP) v1.0
'本代码参考了源码之家的《用ASP编程控制IIS建立WEB站点》和Adsutil.vbs的代码
'还要感谢Envymask和LuoLuo的帮助，谢谢
'小生对ADSI也不是很熟悉，所以代码中有什么错误的，还请大家批评指教，谢谢
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

Function BackupSite() '备份站点模块
	objOutStream.Write "请输入用来备份站点信息的文件名:"
	strBackupFile = objInStream.ReadLine
	intSiteIndex = 1 '站点ID的索引
	Wscript.Echo vbCrLf & "开始备份站点信息……" & vbCrLf
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '建立IIS对象
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断ObjChildObject.Name是不是数字
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name)
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			strServerComment  = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
			strMaxConnections = objIIs.MaxConnections '把IIS虚拟站点的最大连接数赋值给变量MaxConnections
			Set objIIsWeb = objIIs.GetObject("IIsWebVirtualDir","Root")
			strPath = objIIsWeb.Path
			Set objFso = CreateObject("Scripting.FileSystemObject")
			If (objFso.fileexists(strBackupFile)) Then
				Set objBackupFile = objFso.OpenTextFile(strBackupFile,ForAppending,True)
				objBackupFile.WriteLine(strServerComment  &vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath)
				objBackupFile.Close
				Wscript.Echo "正在备份站点 " & strServerComment & "!"
			Else
				Set objBackupFile = objFso.CreateTextFile(strBackupFile,True)
				objBackupFile.WriteLine(strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath)
				objBackupFile.Close
				Wscript.Echo "正在备份站点 " & strServerComment & "!"
			End If
		intSiteIndex = intSiteIndex + 1
		End If
	Next
	Wscript.Echo vbCrLf & "一共备份了 " & intSiteIndex - 1 & "个站点信息!" & vbCrLf
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Set objFso         = Nothing
	Set objBackupFile  = Nothing
End Function

Function RestoreSite() '恢复站点模块
	objOutStream.Write "请输入用来恢复站点信息的文件名:"
	strBackupFile = objInStream.ReadLine
	Wscript.Echo vbCrLf & "开始恢复站点信息……" & vbCrLf
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If Not (objFso.fileexists(strBackupFile)) Then
	Wscrip.Echo "没有找到文件 "&strBackupFile&"!"&vbCrLf&"请确定文件的位置!"
	Wscript.Quit
	End If
	Set objBackupFile = objFso.OpenTextFile(strBackupFile,1)
	intSiteCount = 1
	Do While objBackupFile.AtEndOfStream <> True
		strFileLine = objBackupFile.ReadLine
		arrSiteInfo = split(strFileLine,vbTab)
		intCreateWebStatus = CreateWebServer(arrSiteInfo(0),split(arrSiteInfo(1),","),arrSiteInfo(2),arrSiteInfo(3))
		If intCreateWebStatus = 1 Then
			Wscript.Echo "建立站点 " & arrSiteInfo(0) & " 成功"
			intSiteCount = intSiteCount + 1
		Else
			Wscript.Echo "建立站点 " & arrSiteInfo(0) & " 失败"
		End If
	Loop
	objBackupFile.Close
	Wscript.Echo vbCrLf & "一共恢复了 " & intSiteCount - 1 & " 个站点信息!" & vbCrLf
	Set objBackupFile = Nothing
	Set objFso        = Nothing
End Function


Function ChangeSiteIP() '修改站点IP模块
	'*********************************************
	'获取两个IP，一个是原来的，是个是新的
	'*********************************************
	objOutStream.Write "请输入站点原IP："
	strOldIP = objInStream.ReadLine
	objOutStream.Write "请输入站点新IP："
	strNewIP = objInStream.ReadLine
	'*********************************************
	Wscript.Echo vbCrLf & "开始修改……" & vbCrLf
	Set objW3svc = GetObject("IIS://localhost/w3svc") '建立IIS对象
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断objChildObject.Name是不是数字
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			strServerComment  = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
			Wscript.Echo "正在修改站点  " & strServerComment & "  的IP" '在命令行打印正在修改IP的站点名称
			intBindsIndex = UBound(arrServerBindings) + 1 '由于站点可以绑定多个域名、端口和IP，所以需要判断一共绑定了多少IP
			For intBindsArrID = 1 To intBindsIndex
				If instr(arrServerBindings(intBindsArrID - 1) , strOldIP) Then
					arrServerBindings(intBindsArrID - 1) = Replace(arrServerBindings(intBindsArrID - 1) , strOldIP , strNewIP) '把数组内包含oldip的字符串替换成newip的字符串
				End If
			Next
			objIIs.ServerBindings = arrServerBindings '把iis对象的ServerBindings属性修改为替换过IP的数组
			objIIs.setinfo '使替换过的设置生效
		End If
	Next
	'*********************************************
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Wscript.Echo vbCrLf & "修改完成!"
End Function

Function Change_AppIsolated() '修改应用程序保护模块
	objOutStream.Write "选择你需要设置的(0,低_IIS进程 2,中_共用的 3,高_独立的,默认设置为2):"
	intChoice = objInStream.ReadLine
	If Not IsNumeric(intChoice) Then
		Wscript.Echo "错误，你输入的不是数字!"
		Wscript.Quit
	End If
	intSiteIndex = 1 '站点ID的索引
	Wscript.Echo vbCrLf & "开始修改……" & vbCrLf
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '建立IIS对象
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断取出的objChildObject.Name是不是数字
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
			Wscript.Echo "设置站点 " & strServerComment & " 完成!"
		End If
	Next
	Set objW3svc       = Nothing
	Set objChildObject = Nothing
	Set objIIs         = Nothing
	Set objIIsWeb      = Nothing
	Wscript.Echo vbCrLf & "设置完成!"
End Function

Function Search() '搜索模块
	Wscript.Echo "请选择搜索方式：" & vbCrLf
	Wscript.Echo "1.根据IP查询站点信息"
	Wscript.Echo "2.根据域名查询站点信息"
	Wscript.Echo "3.根据站点绝对路径查询站点信息" & vbCrLf
	objOutStream.Write "请选择你需要的查询方式（1 2 3）："
	strSearchMode = objInStream.ReadLine

	Select Case strSearchMode
	Case "1"
	Use_IP_Search()
	Case "2"
	Use_Domain_Search()
	Case "3"
	Use_Path_Search()
	Case Else
	Wscript.Echo String(30," ") & "输入错误,请重新输入" & vbCrLf
	Search()
	End Select
End Function

Function Use_IP_Search() '根据IP搜索站点信息
	objOutStream.Write "请输入你要搜索的IP(支持模糊搜索)："
	strIpString = objInStream.ReadLine
	Wscript.Echo "开始搜索……" & vbCrLf
	intSiteIndex = 1
	'*********************************************
	Set objW3svc = GetObject("IIS://localhost/w3svc") '建立IIS对象
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断objChildObject.Name是不是数字
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '由于站点可以绑定多个域名、端口和IP，所以需要判断一共绑定了多少IP
			intSearched=0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strIpString) Then
					intSearched=1
					Exit For
				End If
			Next
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				strMaxConnections = objIIs.MaxConnections
				Wscript.Echo "站点描述：" & strServerComment
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '打印出搜索到的站点信息
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "未查找到站点信息!"
	Else
		Wscript.Echo vbCrLf & "共查找到" & intSiteIndex - 1 & "个站点!"
		Wscript.Echo "默认只显示出站点描述，如果您需要详细信息，请把搜索结果存为文件!"
		objOutStream.Write "是否将站点信息存为备份文件？（YES，NO）："
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

Function Use_Domain_Search() '域名搜索模块
	objOutStream.Write "请输入你要搜索的域名支持模糊搜索："
	strDomainString = objInStream.ReadLine
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	Wscript.Echo "开始搜索……" & vbCrLf
	intSiteIndex = 1
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断objChildObject.Name是不是数字
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '由于站点可以绑定多个域名、端口和IP，所以需要判断一共绑定了多少IP
			intSearched=0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strDomainString) Then
					intSearched=1
					Exit For
				End If
			Next
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				strMaxConnections = objIIs.MaxConnections
				Wscript.Echo "站点描述：" & strServerComment
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '打印出搜索到的站点信息
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "未查找到站点信息!"
	Else
		Wscript.Echo vbCrLf & "共查找到" & intSiteIndex - 1 & "个站点!"
		Wscript.Echo "默认只显示出站点描述，如果您需要详细信息，请把搜索结果存为文件!"
		objOutStream.Write "是否将站点信息存为备份文件？（YES，NO）："
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

Function Use_Path_Search() '域名搜索模块
	objOutStream.Write "请输入你要搜索的绝对路径(支持模糊搜索)："
	strPathString = objInStream.ReadLine
	Set objW3svc = GetObject("IIS://LocalHost/W3svc")
	Wscript.Echo "开始搜索……" & vbCrLf
	intSiteIndex = 1
	For Each objChildObject In objW3svc
		If (Err.Number <> 0) Then Exit For
		If IsNumeric(objChildObject.Name) = True Then '判断objChildObject.Name是不是数字
			Set objIIs = objW3svc.GetObject("IIsWebServer", objChildObject.Name) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If instr(objIIsWeb.Path , strPathString) Then
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
				strMaxConnections = objIIs.MaxConnections
				strPath=objIIsWeb.Path
				Wscript.Echo "站点描述：" & strServerComment
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSiteIndex = intSiteIndex + 1
			End If
		End If
	Next
	'Wscript.Echo strSearchedSite '打印出搜索到的站点信息
	If intSiteIndex = 1 Then
		Wscript.Echo vbCrLf & "未查找到站点信息!"
	Else
		Wscript.Echo vbCrLf & "共查找到" & intSiteIndex - 1 & "个站点!"
		Wscript.Echo "默认只显示出站点描述，如果您需要详细信息，请把搜索结果存为文件!"
		objOutStream.Write "是否将站点信息存为备份文件？（YES，NO）："
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
	Wscript.Echo vbCrLf & "请选择批量删除的方式:" & vbCrLf
	Wscript.Echo "1.根据IP删除站点" & vbCrLf
	Wscript.Echo "2.根据域名删除站点" & vbCrlf
	Wscript.Echo "3.根据站点描述删除站点" & vbCrLf
	Wscript.Echo "4.删除所有站点信息" & vbCrLf
	objOutStream.Write "请选择你需要的删除方式（1 2 3 4）："
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
		Wscript.Echo vbCrLf & String(30," ") & "输入错误,请重新输入!"
		DeleteSite()
	End Select
End Function

Function Delete_By_Ip()
	objOutStream.Write "请输入你要删除的站点的IP(支持模糊查询):"
	strIpString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "开始搜索……" & vbCrLf
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
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '由于站点可以绑定多个域名、端口和IP，所以需要判断一共绑定了多少IP
			intSearched = 0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strIpString) Then
					intSearched=1
					Exit For
				End If
			Next
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "删除站点: " & strServerComment & "完成!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "未查找到要删除的站点信息!"
	Else
		Wscript.Echo vbCrLf & "共删除" & intCountDeleteSite & "个站点!"
		objOutStream.Write "是否将已删除的站点信息存为备份文件？（YES，NO）："
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
	objOutStream.Write "请输入你要删除的站点的域名(支持模糊查询):"
	strDomainString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "开始搜索……" & vbCrLf
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
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			intBindsIndex = UBound(arrServerBindings) + 1 '由于站点可以绑定多个域名、端口和IP，所以需要判断一共绑定了多少IP
			intSearched = 0
			For intBindsArrID = 1 To intBindsIndex
				If InStr(arrServerBindings(intBindsArrID - 1) , strDomainString) Then
					intSearched=1
					Exit For
				End If
			Next
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
			If intSearched=1 Then
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "删除站点: " & strServerComment & "完成!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "未查找到要删除的站点信息!"
	Else
		Wscript.Echo vbCrLf & "共删除" & intCountDeleteSite & "个站点!"
		objOutStream.Write "是否将已删除的站点信息存为备份文件？（YES，NO）："
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
	objOutStream.Write "请输入你要删除的站点的描述(支持模糊查询):"
	strCommentString = objInStream.ReadLine
	Wscript.Echo vbCrLf & "开始搜索……" & vbCrLf
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
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '建立IIS虚拟站点对像
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
				arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "删除站点: " & strServerComment & "完成!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intSearched = 0
				intCountDeleteSite = intCountDeleteSite + 1
			End If
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "未查找到要删除的站点信息!"
	Else
		Wscript.Echo vbCrLf & "共删除" & intCountDeleteSite & "个站点!"
		objOutStream.Write "是否将已删除的站点信息存为备份文件？（YES，NO）："
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
	Wscript.Echo vbCrLf & "开始删除……" & vbCrLf
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
			Set objIIs = objW3svc.GetObject("IIsWebServer", arrSiteId(intDeleteId)) '建立IIS虚拟站点对像
			If Err.Number <> 0 Then
				Exit For
				Wscript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
				Err.Clear()
				Wscript.Quit
			End If
			arrServerBindings = objIIs.ServerBindings '把IIS虚拟站点绑定的IP、端口、域名的值放入数组ServerBindings
			Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strServerComment = objIIs.ServerComment '把IIS虚拟站点的名称赋值给变量ServerComment
				strMaxConnections = objIIs.MaxConnections
				Set objIIsWeb=objIIs.GetObject("IIsWebVirtualDir","Root")
				strPath=objIIsWeb.Path
				objW3svc.Delete "IISWebServer" , arrSiteId(intDeleteId)
				Wscript.Echo "删除站点: " & strServerComment & "完成!"
				strSearchedSite = strSearchedSite & strServerComment & vbTab & Join(arrServerBindings,",") & vbTab & strMaxConnections & vbTab & strPath & vbCrLf
				intCountDeleteSite = intCountDeleteSite + 1
	Next
	If intCountDeleteSite = 0 Then
		Wscript.Echo vbCrLf & "未查找到要删除的站点信息!"
	Else
		Wscript.Echo vbCrLf & "共删除" & intCountDeleteSite & "个站点!"
		objOutStream.Write "是否将已删除的站点信息存为备份文件？（YES，NO）："
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


Function SaveToFile(strResult) '保存文件
	objOutStream.Write "请输入保存的文件名："
	strSaveFileName=objInStream.ReadLine
	Set objFso=CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(strSaveFileName) Then
		Wscript.Echo "文件" & strSaveFileName & "已经存在!"
		objOutStream.Write "是否追加写入文件？(Yes,No)"
		strOverWrite = objInStream.ReadLine
		If UCase(strOverWrite) = "YES" Then
			Set objResultFile=objFso.OpenTextFile(strSaveFileName , ForAppending , True)
			objResultFile.Write strResult
			objResultFile.Close
			Wscript.Echo "写入文件" & strSaveFileName & "完成!"
		End If

		If UCase(strOverWrite) = "NO" Then
			objOutStream.Write "请重命名文件："
			strSaveFileName = objInStream.ReadLine
			Set objResultFile = objFso.CreateTextFile(strSaveFileName,True)
			objResultFile.Write strResult
			objResultFile.Close
			Wscript.Echo "保存文件 " & strSaveFileNmae & " 完成!"
		End If
	Else
		Set objResultFile = objFso.CreateTextFile(strSaveFileName , True)
		objResultFile.Write strResult
		objResultFile.Close
		Wscript.Echo "保存文件 " & strSaveFileName & " 完成!"
	End If
	Set objFso = Nothing
End Function


Function CreateWebServer(strServerComment,arrServerBindings,strMaxConnections,strPath)'建立站点
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
	objIIsWeb.AppFriendlyName   = "默认应用程序" 
	objIIsWeb.SetInfo
	Set objW3svc  = Nothing
	Set objIIs    = Nothing
	Set objIIsWeb = Nothing
	CreateWebServer = 1
End Function

Function HELP()
	Wscript.Echo vbCrLf & "                          IIS总管 V1.6 By BlackFox"
	Wscript.Echo "                            http://fox.he100.com" & vbCrLf
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Backup （备份IIS站点信息到一个文本文件！）"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Restroe （从一个文本文件恢复IIS站点信息！）"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs ChangeIP （修改所有IIS站点IP！）"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Change_App （修改所有IIS站点应用程序级别！）"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs Search （根据条件查找站点信息！）"
	Wscript.Echo "Cscript IIS_Manager-1.6.vbs DeleteSite (根据条件批量删除站点)" & vbCrLf
End Function

'想到的新功能
'备份功能增加一个备份站点数据的功能，通过RAR自动打包，默认打包到站点的上级目录，可选择打包的指定的目录
'增加建立隐藏虚拟目录的功能
'批量删除绑定的指定IP
'程序运行时判断IIS是否启动
'恢复站点绝对路径的ACL信息