Function getContent(patrn, strng)
    Set re= New RegExp '����������ʽ��
    re.Pattern = patrn'����ģʽ��
    re.IgnoreCase = True '�����Ƿ������ַ���Сд��
    re.Global = True '����ȫ�ֿ����ԡ�
    Set Matches = re.Execute(strng)'ִ��������
    RetStr0= Matches.Count
    For Each oMatch in Matches'����ƥ�伯�ϡ�
        RetStr=RetStr & oMatch.value
        RetStr1=RetStr1 & oMatch.SubMatches(0)
        RetStr2=RetStr2 & oMatch.length
    Next
    getContent ="Matches.Count:" & RetStr0 & vbnewline & "oMatch.value:" & RetStr & vbnewline & "oMatch.SubMatches(0):" & RetStr1 & vbnewline  & "oMatch.length:" & RetStr2
End Function

msgbox getContent("<title>(.+?)</title>", "<title>1</title><title>22</title><title>333</title><title>4444</title><title>55555</title>")