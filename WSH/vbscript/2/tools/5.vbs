'*********************************************
'*by r05e
'*����GUEST�û�������־
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Function input()
choevt=InputBox("������Ҫ����Guest���ʵ���־��ţ�"&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","ѡ��Ҫ����GUEST���ʵ���־�ļ�", "1", 100, 100)


If choevt<>""  Then
If  IsNumeric(choevt) and choevt<5 and choevt>0  Then

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

sregval = WshShell.RegRead("HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\")
If Err.Number <> 0 Then
    wscript.echo "����������߱�������־���ܣ�����������"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\RestrictGuestAccess","1","REG_DWORD"
wscript.echo "�����ɹ�"
End If
Else wscript.echo "������������������"
wscript.sleep 1000
input()

End If
Else 
wscript.echo "�����Ѿ�ȡ��"

End If

End Function
input()
set WshShell=nothing