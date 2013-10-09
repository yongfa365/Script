'*********************************************
'*转换日志保存的目录
'*********************************************
Set WshShell=CreateObject("Wscript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
Function input()
choevt=InputBox("请输入要更改日志保存文件的日志编号："&vbCrLf&"1-------------Application"&vbCrLf&"2-------------DNS Server"&vbCrLf&"3-------------Security"&vbCrLf&"4-------------System","选择要更改保存文件的日志文件", "1", 100, 100)
evtItem=InputBox("请输入一个绝对路径"," 更改日志保存文件", "%SystemRoot%\system32\config\AppEvent.Evt", 100, 100)



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
wscript.echo "文件不存在或没有用绝对路径，请重新输入"
wscript.sleep 1000
input()
End If
On Error Resume Next

sregval = WshShell.RegRead("HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\File")
If Err.Number <> 0 Then
    wscript.echo "你的主机不具备这项日志功能，请重新输入"
    wscript.sleep 1000
input()

Else 
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\Eventlog\"&evt&"\File",evtItem,"REG_EXPAND_SZ"
wscript.echo "操作成功"
End If
On Error GoTo 0
Else wscript.echo "输入有误，请重新输入"
wscript.sleep 1000
input()

End If

ElseIf choevt="" and evtItem="" Then
wscript.echo "操作已经取消"
Else wscript.echo "输入有误，请重新输入"
wscript.sleep 1000
input()
End If
End Function
input()
set WshShell=nothing
set fso=nothing