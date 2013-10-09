'**********************************************
'*by r05e
'*¹Ø±Õ DirectDraw
'**********************************************
Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\DCI\Timeout","0","REG_DWORD"

set WshShell=nothing