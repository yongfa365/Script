'/*=========================================================================
' * Intro       通过 WMI 改变网卡的IP地址 vbs版
' * FileName    ChangeIP.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365@qq.com
' * FirstWrite  http://www.yongfa365.com/item/Use-WMi-Change-IP-VBS-yongfa365.html
' * LastModify  2007-08-31 14:57:50
' *==========================================================================*/

'strIPAddress = Array("116.90.185.207") '修改后的ip,多个IP可以以","格开,可以写多个
'strSubnetMask = Array("255.255.255.0") '子网掩码,配置同IP
'strGateway = Array("116.90.185.1") '网关
'arrDNSServers = Array("210.51.176.71")'DNS,可以写多个

strIPAddress = Array("192.168.1.38") '修改后的ip,多个IP可以以","格开,可以写多个
strSubnetMask = Array("255.255.255.240") '子网掩码,配置同IP
strGateway = Array("192.168.1.33") '网关
arrDNSServers = Array("202.106.0.20","219.146.0.130")'DNS,可以写多个

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each objNetAdapter in colNetAdapters
    'msgbox colnetadapters.count'看下有几块网卡
    'sip = objNetAdapter.IPAddress '得到原来的ip
    'strIPAddress = sip '保持原来的ip
    strGatewayMetric = Array(1)
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
    errDNS = objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
    If errEnable = 0 Then
        WScript.Echo "IP修改成功"
    Else
        WScript.Echo "IP修改失败"
    End If

    Exit For '只修改第一个网卡的设置
Next