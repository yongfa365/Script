'**************************************
'*by r05e
'*����ϵͳ����汾��
'**************************************
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    Wscript.Echo objOperatingSystem.Caption & " " & objOperatingSystem.Version
Next


