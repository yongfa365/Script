'*********************************************
'*by r05e
'*改变日志文件的大小
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Function input()
choevt=InputBox("请输入要更改大小的日志编号："&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","选择要配置大小的日志文件", "1", 100, 100)
maxsize=InputBox("请输入一个数值（单位为K），比如200或20000,注意：不能超过25000K"," 配置日志文件大小", "512", 100, 100)



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
    wscript.echo "你的主机不具备这项日志功能，请重新输入"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\MaxSize",maxsize,"REG_DWORD"
wscript.echo "操作成功"
End If
Else wscript.echo "输入有误，请重新输入"
wscript.sleep 1000
input()

End If
ElseIf choevt="" and maxsize="" Then
wscript.echo "操作已经取消"
Else wscript.echo "输入有误，请重新输入"
wscript.sleep 1000
input()
End If
End Function
input()
set WshShell=nothing