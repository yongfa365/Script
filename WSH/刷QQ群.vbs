' ˢQQȺ��VBS�ű�      
'���ߣ�С��
'
'ʹ�÷����������´��뱣��ΪQQ.vbs,Ȼ������Ҫ���͵Ķ�����˫��QQ.vbs�Ϳ����Զ�ˢ������
'ע�⣺���������Ϊ�����о�֮�ã�����Ƿ�ʹ�ã� 
'��Һã�����һ��Ⱥ��
Set WshShell= WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "2920"
for i=1 to 3 'Ҫ���͵Ĵ���
WScript.Sleep 1000
WshShell.SendKeys "^v"
WshShell.SendKeys i
WshShell.SendKeys "%s"
next 


WshShell.AppActivate "�� �Եش� ������"
for i=1 to 3 'Ҫ���͵Ĵ���
WScript.Sleep 1000
WshShell.SendKeys "^v"
WshShell.SendKeys i
WshShell.SendKeys "%s"
next 
