'*********************************************
'*by r05e
'*禁止默认共享（FOR SERVER）
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters\AutoShareServer","0","REG_DWORD"

wscript.echo "操作成功"
set WshShell=nothing