'*********************************************
'*ת����־�����Ŀ¼
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
Function input()
choevt=InputBox("������Ҫ������־�����ļ�����־��ţ�"&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","ѡ��Ҫ���ı����ļ�����־�ļ�", "1", 100, 100)
evtItem=InputBox("������һ������·��"," ������־�����ļ�", "%SystemRoot%\system32\config\AppEvent.Evt", 100, 100)



If choevt<>"" and evtItem<>"" Then
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
If (Not fso.FileExists(evtItem)) Then
wscript.echo "�ļ������ڻ�û���þ���·��������������"
wscript.sleep 1000
input()
End If
On Error Resume Next

sregval = WshShell.RegRead("HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\File")
If Err.Number <> 0 Then
    wscript.echo "����������߱�������־���ܣ�����������"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\File",evtItem,"REG_EXPAND_SZ"
wscript.echo "�����ɹ�"
End If
On Error GoTo 0
Else wscript.echo "������������������"
wscript.sleep 1000
input()

End If

ElseIf choevt="" and evtItem="" Then
wscript.echo "�����Ѿ�ȡ��"
Else wscript.echo "������������������"
wscript.sleep 1000
input()
End If
End Function
input()
set WshShell=nothing
set fso=nothing