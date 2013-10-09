'On Error Resume Next	'发生错误就执行下一行

Set WSHShell = CreateObject("WScript.Shell") 
Favorites = WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites") 
Favorites=inputbox("请输入收藏夹路径","第一步：请输入收藏夹路径",Favorites)

if FolderExits(Favorites)=false Then
	msgbox("收藏夹路径不存在，点确定退出")
	WScript.Quit
end if

BadFolder=inputbox("打不开的链接放在哪，默认为：","第二步：打不开的链接放在哪",Favorites&"\有问题的网址\")
CreateFolder(BadFolder)

if FolderExits(BadFolder)=false Then
	msgbox("输入地址不合法，点确定退出")
	WScript.Quit
end if

'msgbox("第三步：开始检测")
Files=""
FilesTree Favorites '遍历
Files=split(Files,"|")

for URLI=1 to Ubound(Files)
    On Error Resume Next
	FileContents=getContents("URL=(.+)", ReadFile(Files(URLI)) ,true)
	If CheckUrl(FileContents(0))=false then
		MoveAFile Files(URLI),BadFolder
	end if
next

msgbox("检测完成")

'判断文件夹是否存在
Function FolderExits(Folder)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Folder) Then
        FolderExits = True
    Else
        FolderExits = False
    End If
End Function

Sub MoveAFile(FileName,sPath)
   Set fso = CreateObject("Scripting.FileSystemObject")
   fso.MoveFile FileName, sPath
End Sub

Function CreateFolder(Folder)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CreateFolder(Folder)
    If Err>0 Then
        Err.Clear
        CreateFolder = False
    Else
        CreateFolder = True
    End If
End Function


Function ReadFile(filename)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(filename, 1, True)
    readfile = cnt.ReadAll
End Function

Function getContents(patrn, strng ,yinyong)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    if yinyong then
	    For i=0 to Matches.count-1
	    	if Matches(i).value<>"" then RetStr = RetStr & Matches(i).SubMatches(0) & "柳永法"
	    Next
	Else
	    For each oMatch in Matches
	    	if oMatch.value<>"" then RetStr = RetStr & oMatch.value & "柳永法"
	    Next
	end if
    getContents = split(RetStr,"柳永法")

End Function

Function CheckUrl(strUrl)
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")
	objHTTP.Open "HEAD", strUrl, FALSE
	objHTTP.Send
	If objHTTP.status=200 Then
		CheckUrl = True
	Else
		CheckUrl = False
	End If
End Function


Function FilesTree(sPath)
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders

    Set oFiles = oFolder.Files
    For Each oFile In oFiles
    	if right(oFile.name,3)="url" then
			Files=Files & "|" &oFile
		end if
    Next
    
    For Each oSubFolder In oSubFolders
        FilesTree oSubFolder.Path '递归
    Next

End Function

