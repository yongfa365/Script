strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\default")
Set colEvents = objWMIService.ExecNotificationQuery("SELECT * FROM RegistryKeyChangeEvent WHERE Hive='HKEY_LOCAL_MACHINE' AND " & "KeyPath='SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run'") 
 
Do
    Set objLatestEvent = colEvents.NextEvent
    MsgBox "在"& Now & "启动项被更改！" ,vbexclamation ,"注册表防火墙"
Loop