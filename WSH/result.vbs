msgbox(String(10,"������"))
msgbox(String(10,65))
msgbox("V"&Space(6)&"V")
msgbox(StrReverse("VBScript"))
MsgBox(TimeValue("4:35:17 PM"))
MsgBox Weekday(Date,vbMonday)
'vbSunday 1 ������ 
'vbMonday 2 ����һ 
'vbTuesday 3 ���ڶ� 
'vbWednesday 4 ������ 
'vbThursday 5 ������ 
'vbFriday 6 ������ 
'vbSaturday 7 ������ 
msgbox WeekdayName(Weekday(Date,vbMonday))

'industr(?:y|ies) ����һ���� 'industry|industries' �����Եı��ʽ��

'Windows (?=95|98|NT|2000)' ��ƥ�� "Windows 2000" �е� "Windows" ��������ƥ�� "Windows 3.1" �е� "Windows"��

��'Windows (?!95|98|NT|2000)' ��ƥ�� "Windows 3.1" �е� "Windows"��������ƥ�� "Windows 2000" �е� "Windows"��

(?:pattern) ƥ�� pattern ������ȡƥ������Ҳ����˵����һ���ǻ�ȡƥ�䣬�����д洢���Ժ�ʹ�á�����ʹ�� "��" �ַ� (|) �����һ��ģʽ�ĸ��������Ǻ����á����磬 'industr(?:y|ies) ����һ���� 'industry|industries' �����Եı��ʽ�� 
(?=pattern) ����Ԥ�飬���κ�ƥ�� pattern ���ַ�����ʼ��ƥ������ַ���������һ���ǻ�ȡƥ�䣬Ҳ����˵����ƥ�䲻��Ҫ��ȡ���Ժ�ʹ�á����磬 'Windows (?=95|98|NT|2000)' ��ƥ�� "Windows 2000" �е� "Windows" ��������ƥ�� "Windows 3.1" �е� "Windows"��Ԥ�鲻�����ַ���Ҳ����˵����һ��ƥ�䷢���������һ��ƥ��֮��������ʼ��һ��ƥ��������������ǴӰ���Ԥ����ַ�֮��ʼ�� 
(?!pattern) ����Ԥ�飬���κβ�ƥ��Negative lookahead matches the search string at any point where a string not matching pattern ���ַ�����ʼ��ƥ������ַ���������һ���ǻ�ȡƥ�䣬Ҳ����˵����ƥ�䲻��Ҫ��ȡ���Ժ�ʹ�á�����'Windows (?!95|98|NT|2000)' ��ƥ�� "Windows 3.1" �е� "Windows"��������ƥ�� "Windows 2000" �е� "Windows"��Ԥ�鲻�����ַ���Ҳ����˵����һ��ƥ�䷢���������һ��ƥ��֮��������ʼ��һ��ƥ��������������ǴӰ���Ԥ����ַ�֮��ʼ  
