'/*=========================================================================
' * Intro       ��������һȦ��������ô�ã���һ���Ƚϲ���ģ��������ߺ���ʱ���˸��Լ��Ľ����ļ�������ļ��ȳ������󣬸о���ˬ�����Ǳ��˵�VBS��MAC�޸Ĵ���㵮����,��ʹ�ù�����������ֲ�������������÷���һ��������������Щ�����Ƚ��ر𣩣����Ҫ������ǰ��MAC���ԣ���ʼ-->�������-->��������-->�����������(һ����"��������")-->��������������-->����..-->�߼�-->ѡ��-->NetworkAddress-->�ұ�ѡ��"������"
' * FileName    ChangeMAC.vbs
' * Author      yongfa365
' * Version     v3.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/ChangeMAC.vbs.html
' * MadeTime    2007-12-09 22:17:58
' * LastModify  2007-12-13 18:35:58
' *==========================================================================*/

On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=true", , 48)
For Each objItem in colItems
    msg = msg & "��ţ�" & objItem.Index & "   MAC��" & objItem.MACAddress & vbCrLf & "������" & objItem.Description & vbCrLf & vbCrLf
Next

idx = InputBox( msg , "1/2��������Ҫ�޸ĵ�MAC�ı��", "1")
If Not IsNumeric(idx) Or Len(idx) = 0 Then
    WScript.Echo "������������˳�"
    Wscript.Quit
End If
MAC = InputBox( "������ָ����MAC��ֵַ��ע��Ӧ����12λ���������ֻ���ĸ�����û��-�����ȷָ�����" , "2/2�������޸ĺ��MAC��ַ", "000000000000")
MAC = Replace(Replace(Replace(MAC, ":", ""), "-", ""), " ", "")
If RegExpTest("[^\da-fA-F]", MAC)>0 Or Len(MAC)<>12 Then
    WScript.Echo "MAC���������˳�"
    Wscript.Quit
End If


idx = Right("00000"&idx, 4)
reg = "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\" & idx
Set WSHShell = CreateObject("WScript.Shell")
WshShell.RegWrite reg & "\NetworkAddress", MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\default" , MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\ParamDesc" , "NetworkAddress" , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\optional" , "1" , "REG_SZ"
'�õ����������ƣ����硰�������� 2��
NetWorkName = WshShell.RegRead("HKLM\SYSTEM\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & WshShell.RegRead(reg & "\NetCfgInstanceId") & "\Connection\Name")

restartNetWork NetWorkName
'WScript.Echo "�޸ĳɹ�"

Function restartNetWork(sConnectionName)
    '��������
    'sConnectionName = "�������� 5" '�ɸĳ���Ҫ���Ƶ��������ƣ���"������������"��
    '��λ����������
    Set shellApp = CreateObject("shell.application")
    Set oControlPanel = shellApp.Namespace(3)
    For Each folderitem in oControlPanel.Items
        If RegExpTest("��������|network",folderitem.Name)<>0 Then
            Set oNetConnections = folderitem.GetFolder
            Exit For
        End If
    Next
    '��λ��Ҫ���������
    For Each folderitem in oNetConnections.Items
        If LCase(folderitem.Name) = LCase(sConnectionName) Then
            Set oLanConnection = folderitem
            Exit For
        End If
    Next
    '��������
    bEnable=True
    ExitOK=False
    Do while 1
        For Each verb in oLanConnection.verbs
	        if RegExpTest("����|disabled",verb.Name)<>0 and bEnable=True then
	                verb.DoIt
	                bEnable=False
	        elseif  RegExpTest("����|enable",verb.Name)<>0 and bEnable=False then
	                verb.DoIt
	                ExitOK=True
	        end if
	        Exit For
        Next
        If ExitOK then exit do
     loop
WScript.Sleep 3000
CreateObject("wscript.shell").popup "�ɹ�",3
End Function


'���������û��ƥ������

Function RegExpTest(patrn, strng)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    RegExpTest = Matches.Count
End Function