'*********************************************
'*by r05e ����Adam�ĸı�����ģ�Adam����ʾ��ʽ���ã�
'*����Ϣ����ʾϵͳ��������
'*********************************************
Set wshshell = CreateObject("WScript.Shell")
For Each EnvirSYSTEM In wshshell.Environment("SYSTEM")
 enOutSYSTEM=enOutSYSTEM&"��ǰ"&EnvirSYSTEM&vbCrlf
Next

For Each EnvirPROCESS In wshshell.Environment("PROCESS")
 enOutPROCESS=enOutPROCESS&"��ǰ"&EnvirPROCESS&vbCrlf
Next

For Each EnvirUSER In wshshell.Environment("USER")
 enOutUSER=enOutUSER&"��ǰ"&EnvirUSER&vbCrlf
Next

For Each EnvirVOLATILE In wshshell.Environment("VOLATILE")
 enOutVOLATILE=enOutVOLATILE&"��ǰ"&EnvirVOLATILE&vbCrlf
Next
WScript.Echo  enOutSYSTEM&enOutPROCESS&enOutUSER&enOutVOLATILE
set wshshell=nothing