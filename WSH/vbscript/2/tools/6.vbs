'*********************************************
'*by r05e
'*��ֹ����IPC����
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Lsa\RestrictAnonymous","1","REG_DWORD"

wscript.echo "�����ɹ�"
set WshShell=nothing