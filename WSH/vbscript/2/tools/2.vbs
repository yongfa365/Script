'*********************************************
'*by r05e
'*�ı���־�ļ��Ĵ�С
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Function input()
choevt=InputBox("������Ҫ���Ĵ�С����־��ţ�"&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","ѡ��Ҫ���ô�С����־�ļ�", "1", 100, 100)
maxsize=InputBox("������һ����ֵ����λΪK��������200��20000,ע�⣺���ܳ���25000K"," ������־�ļ���С", "512", 100, 100)



If choevt<>"" and maxsize<>"" Then
If IsNumeric(maxsize) and IsNumeric(choevt) and choevt<5 and choevt>0 and maxsize>0 and maxsize<25001 Then
maxsize=maxsize*1024
If choevt=1 Then
evt="Application"
ElseIf choevt=2 Then
evt="DNS Server"
ElseIf choevt=3 Then
evt="Security"
ElseIf choevt=4 Then
evt="System"
End If
On Error Resume Next

sregval = WshShell.RegRead("HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\MaxSize")
If Err.Number <> 0 Then
    wscript.echo "����������߱�������־���ܣ�����������"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\MaxSize",maxsize,"REG_DWORD"
wscript.echo "�����ɹ�"
End If
Else wscript.echo "������������������"
wscript.sleep 1000
input()

End If
ElseIf choevt="" and maxsize="" Then
wscript.echo "�����Ѿ�ȡ��"
Else wscript.echo "������������������"
wscript.sleep 1000
input()
End If
End Function
input()
set WshShell=nothing