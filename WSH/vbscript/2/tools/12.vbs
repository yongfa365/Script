'******************************************
'*by r05e
'*��ֹ�����û���½
'******************************************



Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\Network\Logon\MustBeValidated","1","REG_DWORD"
wscript.echo "�����ɹ�"
set WshShell=nothing
