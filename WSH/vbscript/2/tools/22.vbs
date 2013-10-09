'****************************
'*by r05e
'*关机时清除掉页面文件
'****************************
Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown","1","REG_DWORD"

set WshShell=nothing