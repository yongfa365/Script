'*********************************************
'*by r05e
'*禁止匿名IPC连接
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Lsa\RestrictAnonymous","1","REG_DWORD"

wscript.echo "操作成功"
set WshShell=nothing