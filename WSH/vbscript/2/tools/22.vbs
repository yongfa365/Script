'****************************
'*by r05e
'*�ػ�ʱ�����ҳ���ļ�
'****************************
Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\ClearPageFileAtShutdown","1","REG_DWORD"

set WshShell=nothing