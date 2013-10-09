'**************************************
'*by r05e
'*操作系统及其版本号
'**************************************
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    Wscript.Echo objOperatingSystem.Caption & " " & objOperatingSystem.Version
Next


