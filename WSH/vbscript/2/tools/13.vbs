'*********************************************
'*by r05e
'*�ر����445�˿�
'*********************************************

Set WshShell=wscript.CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\NetBT\Parameters\SMBDeviceEnabled",_
                                "0","REG_DWORD"

wscript.echo "�����ɹ�"
set WshShell=nothing