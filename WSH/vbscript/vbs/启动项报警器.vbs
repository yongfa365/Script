strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\default")
Set colEvents = objWMIService.ExecNotificationQuery("SELECT * FROM RegistryKeyChangeEvent WHERE Hive='HKEY_LOCAL_MACHINE' AND " & "KeyPath='SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run'") 
 
Do
    Set objLatestEvent = colEvents.NextEvent
    MsgBox "��"& Now & "��������ģ�" ,vbexclamation ,"ע������ǽ"
Loop