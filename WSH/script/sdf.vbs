Dim cm 
Set cm= CreateObject("CDO.Message") 
'�������� 
cm.From="cbd@cbdcn.com" 
'���÷����˵����� 
cm.To="yongfa365@qq.com" 
'���������˵����� 
cm.Subject="�ҷ���һ���������߶���RSS����վ���Ժ����ٰ�װʲô����ˡ�" 
'�趨�ʼ������� 
'cm.TextBody="http://www.knowsky.com/rss/ " 
'������ʹ����ͨ���ı���ʽ�����ʼ���ֻ�������֣�����֧��html���������ﲻ�� 

cm.HtmlBody="ok" 

'��������㹹����html���ģ������������ʼ��ͱ�ֻ�����ֵĺÿ����ˡ���Ҫ˵�㲻��html�� 

'cm.AddAttachment HTML & vbcrlf & "test.zip"
'�������Ҫ���͸����Ļ���������ķ������ļ����ӽ�ȥ�� 

cm.Send 
'���Ȼ��ִ�з����� 
Set cm=Nothing 
'���ͳɹ���ʱ�ͷŶ��� 

msgbox("�����ʼ��ɹ���") 
