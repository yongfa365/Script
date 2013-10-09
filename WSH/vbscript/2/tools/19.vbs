'****************************************
'*by r05e
'*关机
'****************************************
Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem")
Function shut()
setTime=InputBox("请输入多少时间后关机（单位：秒）"," 定时关机", "", 100, 100)
If setTime<>"" Then
IF IsNumeric(setTime) Then
timeS=setTime*1000
wscript.sleep timeS


For Each objOperatingSystem in colOperatingSystems
    ObjOperatingSystem.Win32Shutdown(1)
Next
Else wscript.echo "输入错误,请重试"
shut()
End If
Else wscript.echo "操作取消"
End If
End function
shut()