'/*=========================================================================
' * Intro       ͨ�� WMI �ı�������IP��ַ vbs��
' * FileName    ChangeIP.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365@qq.com
' * FirstWrite  http://www.yongfa365.com/item/Use-WMi-Change-IP-VBS-yongfa365.html
' * LastModify  2007-08-31 14:57:50
' *==========================================================================*/

'strIPAddress = Array("116.90.185.207") '�޸ĺ��ip,���IP������","��,����д���
'strSubnetMask = Array("255.255.255.0") '��������,����ͬIP
'strGateway = Array("116.90.185.1") '����
'arrDNSServers = Array("210.51.176.71")'DNS,����д���

strIPAddress = Array("192.168.1.38") '�޸ĺ��ip,���IP������","��,����д���
strSubnetMask = Array("255.255.255.240") '��������,����ͬIP
strGateway = Array("192.168.1.33") '����
arrDNSServers = Array("202.106.0.20","219.146.0.130")'DNS,����д���

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each objNetAdapter in colNetAdapters
    'msgbox colnetadapters.count'�����м�������
    'sip = objNetAdapter.IPAddress '�õ�ԭ����ip
    'strIPAddress = sip '����ԭ����ip
    strGatewayMetric = Array(1)
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
    errDNS = objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
    If errEnable = 0 Then
        WScript.Echo "IP�޸ĳɹ�"
    Else
        WScript.Echo "IP�޸�ʧ��"
    End If

    Exit For 'ֻ�޸ĵ�һ������������
Next