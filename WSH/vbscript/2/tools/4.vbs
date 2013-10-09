'*********************************************
'*by r05e
'*cmdhere安装，使你在任何目录下直接进入改目录的CMD状态
'*********************************************

Set WshShell=CreateObject("Wscript.Shell")
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\",""
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\command\",""
WshShell.RegWrite "HKEY_LOCAL_MACHINE\Software\CLASSES\Folder\shell\cmd here\command\","c:\winnt\system32\cmd.exe /K CD %1","REG_SZ"
wscript.echo "操作成功"
set WshShell=nothing