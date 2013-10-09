'/*=========================================================================
' * Intro       通过腾讯临时会话代码实现：QQ强行聊天工具
' * FileName    QQ-Chat.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/blog/item/QQ-QiangXingLiaoTianGongJu.htm
' * LastModify  2007-08-31 15:11:50
' *==========================================================================*/

On Error Resume Next
num = InputBox("请输入对方QQ号码:", "强行聊天工具(单人)") '输入对方QQ号码
If num = "" Then WScript.Quit '如果输入的变量为空则结束程序
url = "tencent://message/?uin=" + num + "&Site=1&Menu=yes" '腾讯临时会话代码
Set obj = WScript.CreateObject("InternetExplorer.Application") '创建IE浏览器对象
obj.Visible = False '设置IE浏览器为不可见
obj.Navigate(url) '打开IE浏览器 ，效果为打开QQ临时会话窗口

