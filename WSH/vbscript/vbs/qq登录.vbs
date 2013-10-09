Dim strPrgpth
strPrgpth = "D:\Program Files\Tencent\QQ\qq.exe"	'如果QQ安装路径不同，在此处修改
Set wshshell = CreateObject("wscript.shell")
Set oexec = wshshell.exec(strPrgpth)
wscript.sleep 3000	'如果脚本无法工作，可以适当延长这部分的时间值
wshshell.AppActivate "QQ用户登录"
wshshell.Sendkeys "{TAB}"
wshshell.Sendkeys "12345678"		'换成你自己的QQ号
wscript.Sleep 1000
wshshell.Sendkeys "{TAB}"
wscript.Sleep 1000
wshshell.Sendkeys "password"		'换成你的QQ号密码
wscript.Sleep 1000
wshshell.Sendkeys "{ENTER}"
wscript.Quit