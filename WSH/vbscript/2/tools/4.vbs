'*********************************************
'*by r05e
'*cmdhere��װ��ʹ�����κ�Ŀ¼��ֱ�ӽ����Ŀ¼��CMD״̬
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\",""
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\command\",""
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\command\","c:\winnt\system32\cmd.exe /K CD %1","REG_SZ"
wscript.echo "�����ɹ�"
set WshShell=nothing