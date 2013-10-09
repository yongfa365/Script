'/*=========================================================================
' * Intro       通过 WMI 改变网卡的IP地址 vbs版
' * FileName    ChangeIP.vbs
' * Author      yongfa365
' * Version     v2.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365@qq.com
' * FirstWrite  http://www.yongfa365.com/item/Use-WMi-Change-IP-VBS-yongfa365.html
' * LastModify  2007-08-31 14:57:50
' *==========================================================================*/

strIPAddress = Array("61.237.236.15") '修改后的ip,多个IP可以以","格开,可以写多个
strSubnetMask = Array("255.255.255.224") '子网掩码,配置同IP
strGateway = Array("61.237.236.1") '网关
arrDNSServers = Array("61.233.9.9")'DNS,可以写多个

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each objNetAdapter in colNetAdapters
    sip = Join(objNetAdapter.IPAddress, ",")
	    strGatewayMetric = Array(1)
	    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
	    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
	    errDNS = objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
Next