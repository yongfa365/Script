'*********************************************
'*by r05e
'*��ֹ��ʾ�ϴε�½�û���
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Winlogon\DontDisplayLastUserName","1","REG_DWORD"
wscript.echo "�����ɹ�"
set WshShell=nothing