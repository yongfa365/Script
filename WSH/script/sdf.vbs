Dim cm 
Set cm= CreateObject("CDO.Message") 
'创建对象 
cm.From="cbd@cbdcn.com" 
'设置发信人的邮箱 
cm.To="yongfa365@qq.com" 
'设置收信人的邮箱 
cm.Subject="我发现一个可以在线订阅RSS的网站，以后不用再安装什么软件了。" 
'设定邮件的主题 
'cm.TextBody="http://www.knowsky.com/rss/ " 
'上面是使用普通的文本格式发送邮件，只能是文字，不能支持html，所以这里不用 

cm.HtmlBody="ok" 

'上面就是你构建的html正文，这样发出的邮件就比只有文字的好看多了。不要说你不会html吧 

'cm.AddAttachment HTML & vbcrlf & "test.zip"
'如果有需要发送附件的话就用上面的方法把文件附加进去。 

cm.Send 
'最后当然是执行发送了 
Set cm=Nothing 
'发送成功后即时释放对象 

msgbox("发送邮件成功。") 
