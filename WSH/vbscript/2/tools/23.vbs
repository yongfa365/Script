'********************************
'*by r05e
'*�����ű�����Ϣ��ʾ��
'********************************
Function loginmsg()
msg=InputBox("������Ҫ�������Ϣ���ݣ�������ÿ���û�LOGINʱ����","������Ϣ", "ע�ⲻҪ���ʷǷ�������Դ", 100, 100)
title=InputBox("��������Ϣ�ı���","�������", "����Ա��ʾ", 100, 100)
prompt=InputBox("������ϵͳ��λ��","����ϵͳ��λ��", "C", 100, 100)
If msg<>"" and title<>"" and prompt<>"" Then
If Len(prompt)=1 and Asc(Lcase(prompt))>96 and Asc(Lcase(prompt))<123 Then
set fso=createobject("scripting.FileSystemObject")
set crVBS=fso.createtextfile(prompt&":\Documents and Settings\All Users\����ʼ���˵�\����\����\msg.vbs",True)

crVBS.writeline("MsgBox "&Chr(34)&msg&Chr(34)&",64,"&Chr(34)&title&Chr(34))
crVBS.close
wscript.echo "�����ɹ�"
End If
ElseIF msg="" and title="" and prompt="" Then
wscript.echo "�����Ѿ�ȡ��"
Else wscript.echo "�������������"
loginmsg()
End If
set fso=nothing
set crVBS=nothing
End Function
loginmsg()