'*********************************************
'*by r05e
'*��ֹĬ�Ϲ���(FOR professional)
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters\AutoShareWks","0","REG_DWORD"
wscript.echo "�����ɹ�"
set WshShell=nothing