'**********************************************
'*by r05e
'*�ر� DirectDraw
'**********************************************
Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\DCI\Timeout","0","REG_DWORD"

set WshShell=nothing