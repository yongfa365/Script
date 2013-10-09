'*********************************************
'*by r05e
'*关闭你的445端口
'*********************************************

Set WshShell=wscript.CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\NetBT\Parameters\SMBDeviceEnabled",_
                                "0","REG_DWORD"

wscript.echo "操作成功"
set WshShell=nothing