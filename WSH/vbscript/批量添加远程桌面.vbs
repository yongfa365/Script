'/*=========================================================================
' * Intro       �������ж�̨�������Ҷ�ʹ��ϵͳ�Դ���Զ���������ӣ�������ϵͳ�󻹵�һ��һ���üӰ��졣Ч�ʷǳ��͵�����
' * FileName    �������Զ������.vbs
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
'ip(0)=Array("������Զ������IP","������Զ�������û���","������Զ����������")
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
