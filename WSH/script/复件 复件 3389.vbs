on error resume next
Function DateFormat(str)
    Set re = New RegExp '����������ʽ��
    re.Pattern = "(\d{4})[-/��]*(\d{2})[-/��]*(\d{2})[-/��]*.*(\d{2}):(\d{2}).*"'����ģʽ��
    re.IgnoreCase = True '�����Ƿ������ַ���Сд��
    re.Global = True '����ȫ�ֿ����ԡ�
    Set Matches = re.Execute(str)'ִ��������
    
	with Matches(0)
		a0=.SubMatches(0)
		a1=.SubMatches(1)
		a2=.SubMatches(2)
		a3=.SubMatches(3)
		a4=.SubMatches(4)

	end with
msgbox a0
msgbox a1
msgbox a2
msgbox a3
msgbox a4


end Function

DateFormat "2003-12-12 12:35"