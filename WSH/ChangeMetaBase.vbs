'/*=========================================================================
' * Intro       让IIS建立的站点默认是.net 2.0的,而不是.net 1.1的,没有使用WMI,所以在操作前先得停止IIS相关服务
' * FileName    ChangeMetaBaseScriptMaps.vbs
' * Author      yongfa365
' * Version     v2.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/ChangeMetaBaseScriptMaps.vbs.html
' * MadeTime    2008-07-24 17:38:20
' * LastModify  2008-08-11 20:36:00   上次发时出错了，现在修正一下
'另：我发的文章在网上搜索时只找到了转载我文章的人的地址，而我的地址根本就没有，很是郁闷
' *==========================================================================*/

WScript.Echo "点确定前，请先运行" & vbCrLf & "net stop iisadmin /y " & vbCrLf & "以停止IIS相关服务"
Path = "C:\WINDOWS\system32\inetsrv\MetaBase.xml"
Node = "//configuration/MBProperty/IIsWebService"
Set XmlDom = CreateObject("msxml2.domdocument")
XmlDom.async = False
XmlDom.load(Path)
ScriptMaps = XmlDom.selectSingleNode(Node).getAttribute("ScriptMaps")
ScriptMaps = Replace(ScriptMaps, "v1.1.4322", "v2.0.50727")
XmlDom.selectSingleNode(Node).setAttribute("ScriptMaps") = ScriptMaps
XmlDom.Save(Path)
WScript.Echo "OK,请运行" & vbCrLf & "iisreset" & vbCrLf & "重启IIS相关服务"