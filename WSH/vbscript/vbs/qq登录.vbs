Dim strPrgpth
strPrgpth = "D:\Program Files\Tencent\QQ\qq.exe"	'���QQ��װ·����ͬ���ڴ˴��޸�
Set wshshell = CreateObject("wscript.shell")
Set oexec = wshshell.exec(strPrgpth)
wscript.sleep 3000	'����ű��޷������������ʵ��ӳ��ⲿ�ֵ�ʱ��ֵ
wshshell.AppActivate "QQ�û���¼"
wshshell.Sendkeys "{TAB}"
wshshell.Sendkeys "12345678"		'�������Լ���QQ��
wscript.Sleep 1000
wshshell.Sendkeys "{TAB}"
wscript.Sleep 1000
wshshell.Sendkeys "password"		'�������QQ������
wscript.Sleep 1000
wshshell.Sendkeys "{ENTER}"
wscript.Quit