'/*=========================================================================
' * Intro       Windowsϵͳ����ǽ�˿���ӣ�һ�������ǵ�öಽȻ����Ӻö�β��������Ӳ�������ʱ�����������һЩ�˿ڣ���δ������ʵ�������Ĺ��ܣ�һ���������Ԥ���趨�õķ���ǽ�˿ڡ�
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

	objPort.Name = strName '����
	objPort.Port = iPort '�˿ں�
	objPort.Protocol = iProtocol 'TCP--> 6,UDP-->17
	objPort.Scope = iScope '��Χall-->0 ,���ҵ�����-->1
	objPort.Enabled = bEnabled '�Ƿ��� True or False

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
AddFirewallPort "�Զ���Զ������˿�",12345,6,0,True
AddFirewallPort "QQ",8000,17,0,True
