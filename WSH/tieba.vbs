'baidu����ˢ���ӵ�vbs�ű� by geniux2008-04-04 04:09 P.M.
set ie=createobject("internetexplorer.application")
ie.toolbar=0
ie.visible=1
ie.navigate "http://post.baidu.com/f?kw=%BA%AA%B5%A6%B6%FE%D6%D0"
wscript.sleep 2000
for i=1 to 100
ie.document.post.ti.value="��ˮ�������"
ie.document.post.co.click
ie.document.post.co.value="baidu ���ɹ�ˮ�������."+vbcrlf+"-----By Geniux"
ie.document.post.submit3.click
wscript.sleep 30*1000
next
msgbox "End" 
