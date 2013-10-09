<%  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
' VBScript 使用 xmldom 检测/创建/读取/更改 XML 文件数据 实例 By shawl.qiu 
'   http://blog.csdn.net/btbtd 
' shawl.qiuATgmail.com 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
        path=server.MapPath("config.xml") '要操作的文件路径 
    set objxml=createObject("microsoft.xmldom") '创建 XMLDOM 对象 
        objxml.load(path) '加载 XML 文件 
        if objxml.parseError<>0 then '//检测文件是否存在 0为存在;不等于0为不存在 
            '如果文件不存在, 执行以下操作 
             
            objxml.appendChild(objxml.createElement("root")) '//创建根节点 
             
            '-' //创建根节点 子节点1 Child1 
            objxml.selectSingleNode("//root").appendChild(objxml.createElement("Child1")) 
             
            '// 创建子节点1 Child1 的 子节点 Children1, 并添加文本 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children1"))_ 
            .text="<h1>Child1 Children1 text</h1>" 
             
            '-'' //创建子节点1 Child1 的子节点 Children2, 并添加 CDATA 文本 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children2")) 
            objxml.selectSingleNode("//root/Child1/Children2").appendChild(objxml.createCDATASection _  
            ("<h2>Child1 Children2 CDATASection</h2>")) 
             
            '-''' //创建子节点1 的子节点 Children3 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children3")) 
             
            ''// 创建根节点 子节点2 
            objxml.selectSingleNode("//root").appendChild(objxml.createElement("Child2")) 
             
            '--- // 添加 XML 文件 的 结构标识 
            objxml.insertBefore objxml.createProcessingInstruction("xml","version=""1.0""" &_  
            " encoding=""utf-8"""), objxml.firstChild 
             
            '//保存文件 
            objxml.save(path) 
        end if 
            '''-' // 读出根节点及其子节点的所有文本, 也就是读出 XML 文档的全部文本 
            response.write objxml.selectSingleNode("//root").text&"<br/>" 
             
            '''-'' // 读出子节点 Child1 及其子节点的所有文本数据 
            response.write objxml.selectSingleNode("//root/Child1").text&"<br/>" 
             
            ''''-' // 读出子节点 Child1 的子节点 Children2 的文本数据 
            response.write objxml.selectSingleNode("//root/Child1/Children2").text&"<br/>" 
             
            ''''-'' // 为 节点1 的子节点 Children3 赋文本值, 并读出文本数据 
            objxml.selectSingleNode("//root/Child1/Children3").text="Root Child1 Children3 text" 
            response.write objxml.selectSingleNode("//root/Child1/Children3").text&"<br/>" 
             
            ''''-'' // 清空子节点 Child1 的子节点 Children2 的文本数据, 并添加 CDATA 数据, 再读出文本数据 
            objxml.selectSinglenode("//root/Child1/Children2").text="" 
            objxml.selectSinglenode("//root/Child1/Children2").appendChild(objxml.createCDATASection("<h3>Child1 Children2 new CDATASection</h3>")) 
            response.write objxml.selectSinglenode("//root/Child1/Children2").text&"<br/>" 
             
            ''--- //保存数据 
            objxml.save(path) 
    set objxml=nothing 
     
    '//小提示: 关于 是否使用 CDATASection 
    '如果只使用 Response.write 读出 XML 文件数据, 大概不需要使用 CDATASECTION, Response.write 自动 转换 
    '特殊的 ASCII 字符为 HTML 字符 
    '如果使用 FSO 读取 XML 数据, 有点牛头不对马嘴的味道 
%> 
