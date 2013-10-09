'==========================================================================
'
' 注销/重起/关闭本地Windows NT/2000 计算机。基本思路如下：
'
' Win32ShutDown(flag)中flag的参数:
' 0 注销
' 0 + 4 强制注销
' 1 关机
' 1 + 4 强制关机
' 2 重起
' 2 + 4 强制重起
' 8 关闭电源
' 8 + 4 强制关闭电源
'
'==========================================================================

ShutDown

Sub ShutDown()
    Dim Connection, WQL, SystemClass, System
    Set Connection = GetObject("winmgmts:root\cimv2")
    WQL = "Select Name From Win32_OperatingSystem"
    Set SystemClass = Connection.ExecQuery(WQL)
    For Each System In SystemClass
        System.Win32ShutDown (2 + 4)
    Next
End Sub
