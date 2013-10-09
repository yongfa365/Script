'/*=========================================================================
' * Intro       Windows系统防火墙端口添加，一般我们是点好多步然后添加好多次才能完成添加操作，有时还会忘记添加一些端口，这段代码可以实现这样的功能：一键添加所有预先设定好的防火墙端口。
' * FileName    AddFirewallPort.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/AddFirewallPort.vbs.html
' * MadeTime    2007-10-13 22:10:49
' * LastModify  2007-10-13 22:10:49
' *==========================================================================*/

ON ERROR RESUME NEXT
Function AddFirewallPort(strName,iPort,iProtocol,iScope,bEnabled)
	Set objFirewall = CreateObject("HNetCfg.FwMgr")
	Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
	Set objPort = CreateObject("HNetCfg.FwOpenPort")

	objPort.Name = strName '名称
	objPort.Port = iPort '端口号
	objPort.Protocol = iProtocol 'TCP--> 6,UDP-->17
	objPort.Scope = iScope '范围all-->0 ,仅我的子网-->1
	objPort.Enabled = bEnabled '是否开启 True or False

	Set colPorts = objPolicy.GloballyOpenPorts
	errReturn = colPorts.Add(objPort)
End Function

AddFirewallPort "WEB 80",80,6,0,True
AddFirewallPort "Imail",8383,6,0,True
AddFirewallPort "Serv-U",21,6,0,True
AddFirewallPort "MSSQL",1433,6,0,True
AddFirewallPort "PASV 5000",5000,6,0,True
AddFirewallPort "PASV 5001",5001,6,0,True
AddFirewallPort "PASV 5002",5002,6,0,True
AddFirewallPort "PASV 5003",5003,6,0,True
AddFirewallPort "自定义远程桌面端口",12345,6,0,True
AddFirewallPort "QQ",8000,17,0,True
