'****************************************
'*by r05e
'*�ػ�
'****************************************
Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem")
Function shut()
setTime=InputBox("���������ʱ���ػ�����λ���룩"," ��ʱ�ػ�", "", 100, 100)
If setTime<>"" Then
IF IsNumeric(setTime) Then
timeS=setTime*1000
wscript.sleep timeS


For Each objOperatingSystem in colOperatingSystems
    ObjOperatingSystem.Win32Shutdown(1)
Next
Else wscript.echo "�������,������"
shut()
End If
Else wscript.echo "����ȡ��"
End If
End function
shut()