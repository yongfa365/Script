'*********************************************
'*by r05e
'*��ֹĬ�Ϲ���FOR SERVER��
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters\AutoShareServer","0","REG_DWORD"

wscript.echo "�����ɹ�"
set WshShell=nothing