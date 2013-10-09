'/*=========================================================================
' * Intro       用XMLDOM分析QQ签名文档
' * FileName    QQ_QianMing.vbs
' * Author      yongfa365
' * Version     v2.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/QQ_QianMing.vbs.html
' * MadeTime    2008-01-22 20:55:25
' * LastModify  2008-02-22 20:55:25
' *==========================================================================*/

Set Doc = CreateObject("Microsoft.XMLDOM")
Doc.async = False
Doc.load("http://e.cnc.qzone.qq.com/cgi-bin/cgi_emotion_indexcount.cgi?uin=64049027")
Set root = Doc.documentElement
Set node = root.childNodes.nextNode()
Wscript.Echo "共" & node.text & "条签名信息"

Set Doc = CreateObject("Microsoft.XMLDOM")
Doc.async = False
Doc.load("http://e.cnc.qzone.qq.com/cgi-bin/cgi_emotion_indexlist.cgi?uin=64049027&emotionarchive=-1")
Set root = Doc.documentElement
Wscript.Echo "XML根结点名字是：" & root.nodeName
Set node = root.childNodes.nextNode()

For nodei = 0 To node.childNodes.Length -1
    Set NowNode = node.childNodes(nodei)
    msg = msg & vbCrLf & "id" & ":" & NowNode.Attributes.getNamedItem("id").text
    '    msg = msg & vbCrLf & "id" & ":" & NowNode.getAttribute("id")
    msg = msg & vbCrLf & "title" & ":" & NowNode.selectSingleNode("title").text
    msg = msg & vbCrLf & "pubDate" & ":" & NowNode.selectSingleNode("pubDate").text
    '    msg = msg & vbCrLf &  NowNode.childNodes(0).nodeName & ":" & NowNode.childNodes(0).text
    '    msg = msg & vbCrLf &  NowNode.childNodes(1).nodeName & ":" & NowNode.childNodes(1).text
Next
Wscript.Echo msg
