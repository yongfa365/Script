' 刷QQ群的VBS脚本      
'作者：小竹
'
'使用方法：讲以下代码保存为QQ.vbs,然后复制你要发送的东西，双击QQ.vbs就可以自动刷屏拉！
'注意：本代码仅作为技术研究之用，请勿非法使用！ 
'大家好，测试一下群发
Set WshShell= WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "2920"
for i=1 to 3 '要发送的次数
WScript.Sleep 1000
WshShell.SendKeys "^v"
WshShell.SendKeys i
WshShell.SendKeys "%s"
next 


WshShell.AppActivate "与 吃地带 聊天中"
for i=1 to 3 '要发送的次数
WScript.Sleep 1000
WshShell.SendKeys "^v"
WshShell.SendKeys i
WshShell.SendKeys "%s"
next 
