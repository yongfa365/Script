'*********************************************
'*by r05e
'*限制GUEST用户访问日志
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Function input()
choevt=InputBox("请输入要限制Guest访问的日志编号："&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","选择要限制GUEST访问的日志文件", "1", 100, 100)


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
    wscript.echo "你的主机不具备这项日志功能，请重新输入"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\RestrictGuestAccess","1","REG_DWORD"
wscript.echo "操作成功"
End If
Else wscript.echo "输入有误，请重新输入"
wscript.sleep 1000
input()

End If
Else 
wscript.echo "操作已经取消"

End If

End Function
input()
set WshShell=nothing