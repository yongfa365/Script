'/*=========================================================================
' * Intro       ��IIS������վ��Ĭ����.net 2.0��,������.net 1.1��,û��ʹ��WMI,�����ڲ���ǰ�ȵ�ֹͣIIS��ط���
' * FileName    ChangeMetaBaseScriptMaps.vbs
' * Author      yongfa365
' * Version     v2.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/ChangeMetaBaseScriptMaps.vbs.html
' * MadeTime    2008-07-24 17:38:20
' * LastModify  2008-08-11 20:36:00   �ϴη�ʱ�����ˣ���������һ��
'���ҷ�����������������ʱֻ�ҵ���ת�������µ��˵ĵ�ַ�����ҵĵ�ַ������û�У���������
' *==========================================================================*/

WScript.Echo "��ȷ��ǰ����������" & vbCrLf & "net stop iisadmin /y " & vbCrLf & "��ֹͣIIS��ط���"
Path = "C:\WINDOWS\system32\inetsrv\MetaBase.xml"
Node = "//configuration/MBProperty/IIsWebService"
Set XmlDom = CreateObject("msxml2.domdocument")
XmlDom.async = False
XmlDom.load(Path)
ScriptMaps = XmlDom.selectSingleNode(Node).getAttribute("ScriptMaps")
ScriptMaps = Replace(ScriptMaps, "v1.1.4322", "v2.0.50727")
XmlDom.selectSingleNode(Node).setAttribute("ScriptMaps") = ScriptMaps
XmlDom.Save(Path)
WScript.Echo "OK,������" & vbCrLf & "iisreset" & vbCrLf & "����IIS��ط���"