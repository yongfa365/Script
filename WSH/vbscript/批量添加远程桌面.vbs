'/*=========================================================================
' * Intro       适用与有多台服务器且都使用系统自带的远程桌面连接，而重做系统后还得一个一个得加半天。效率非常低的问题
' * FileName    批量添加远程桌面.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/PiLiangTianJiaYuanChengZhuoMian.vbs.html
' * MadeTime    2007-11-29 00:46:30
' * LastModify  2007-11-29 00:46:30
' *==========================================================================*/

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.run("%SystemRoot%\system32\tsmmc.msc /s") 
WScript.Sleep 3000

Dim ip(100)
'ip(0)=Array("服务器远程桌面IP","服务器远程桌面用户名","服务器远程桌面密码")
ip(0)=Array("100.110.111.112","UserName0","PassWord0")
ip(1)=Array("111.222.111.121:3389","UserName1","PassWord1")
ip(2)=Array("111.222.111.211:1234","UserName2","PassWord2")
ip(3)=Array("111.222.111.222","UserName3","PassWord3")

For i=0 To 3
	WshShell.SendKeys "+{F10}"
	WshShell.SendKeys "A"
	WshShell.SendKeys ip(i)(0)
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys ip(i)(1)
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys ip(i)(2)
	WshShell.SendKeys "{ENTER}"
Next
