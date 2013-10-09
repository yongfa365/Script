'******************************************
'*by r05e
'*防止匿名用户登陆
'******************************************



Set WshShell=CreateObject("Wscript.Shell")


WshShell.RegWrite "HKLM\Network\Logon\MustBeValidated","1","REG_DWORD"
wscript.echo "操作成功"
set WshShell=nothing
