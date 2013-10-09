'baidu贴吧刷帖子的vbs脚本 by geniux2008-04-04 04:09 P.M.
set ie=createobject("internetexplorer.application")
ie.toolbar=0
ie.visible=1
ie.navigate "http://post.baidu.com/f?kw=%BA%AA%B5%A6%B6%FE%D6%D0"
wscript.sleep 2000
for i=1 to 100
ie.document.post.ti.value="灌水程序测试"
ie.document.post.co.click
ie.document.post.co.value="baidu 贴吧灌水程序测试."+vbcrlf+"-----By Geniux"
ie.document.post.submit3.click
wscript.sleep 30*1000
next
msgbox "End" 
