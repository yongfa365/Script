'*********************************************
'*by r05e
'*禁止显示上次登陆用户名
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")

WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Winlogon\DontDisplayLastUserName","1","REG_DWORD"
wscript.echo "操作成功"
set WshShell=nothing