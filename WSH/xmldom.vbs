<%  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
' VBScript ʹ�� xmldom ���/����/��ȡ/���� XML �ļ����� ʵ�� By shawl.qiu 
'   http://blog.csdn.net/btbtd 
' shawl.qiuATgmail.com 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
        path=server.MapPath("config.xml") 'Ҫ�������ļ�·�� 
    set objxml=createObject("microsoft.xmldom") '���� XMLDOM ���� 
        objxml.load(path) '���� XML �ļ� 
        if objxml.parseError<>0 then '//����ļ��Ƿ���� 0Ϊ����;������0Ϊ������ 
            '����ļ�������, ִ�����²��� 
             
            objxml.appendChild(objxml.createElement("root")) '//�������ڵ� 
             
            '-' //�������ڵ� �ӽڵ�1 Child1 
            objxml.selectSingleNode("//root").appendChild(objxml.createElement("Child1")) 
             
            '// �����ӽڵ�1 Child1 �� �ӽڵ� Children1, ������ı� 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children1"))_ 
            .text="<h1>Child1 Children1 text</h1>" 
             
            '-'' //�����ӽڵ�1 Child1 ���ӽڵ� Children2, ����� CDATA �ı� 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children2")) 
            objxml.selectSingleNode("//root/Child1/Children2").appendChild(objxml.createCDATASection _  
            ("<h2>Child1 Children2 CDATASection</h2>")) 
             
            '-''' //�����ӽڵ�1 ���ӽڵ� Children3 
            objxml.selectSingleNode("//root/Child1").appendChild(objxml.createElement("Children3")) 
             
            ''// �������ڵ� �ӽڵ�2 
            objxml.selectSingleNode("//root").appendChild(objxml.createElement("Child2")) 
             
            '--- // ��� XML �ļ� �� �ṹ��ʶ 
            objxml.insertBefore objxml.createProcessingInstruction("xml","version=""1.0""" &_  
            " encoding=""utf-8"""), objxml.firstChild 
             
            '//�����ļ� 
            objxml.save(path) 
        end if 
            '''-' // �������ڵ㼰���ӽڵ�������ı�, Ҳ���Ƕ��� XML �ĵ���ȫ���ı� 
            response.write objxml.selectSingleNode("//root").text&"<br/>" 
             
            '''-'' // �����ӽڵ� Child1 �����ӽڵ�������ı����� 
            response.write objxml.selectSingleNode("//root/Child1").text&"<br/>" 
             
            ''''-' // �����ӽڵ� Child1 ���ӽڵ� Children2 ���ı����� 
            response.write objxml.selectSingleNode("//root/Child1/Children2").text&"<br/>" 
             
            ''''-'' // Ϊ �ڵ�1 ���ӽڵ� Children3 ���ı�ֵ, �������ı����� 
            objxml.selectSingleNode("//root/Child1/Children3").text="Root Child1 Children3 text" 
            response.write objxml.selectSingleNode("//root/Child1/Children3").text&"<br/>" 
             
            ''''-'' // ����ӽڵ� Child1 ���ӽڵ� Children2 ���ı�����, ����� CDATA ����, �ٶ����ı����� 
            objxml.selectSinglenode("//root/Child1/Children2").text="" 
            objxml.selectSinglenode("//root/Child1/Children2").appendChild(objxml.createCDATASection("<h3>Child1 Children2 new CDATASection</h3>")) 
            response.write objxml.selectSinglenode("//root/Child1/Children2").text&"<br/>" 
             
            ''--- //�������� 
            objxml.save(path) 
    set objxml=nothing 
     
    '//С��ʾ: ���� �Ƿ�ʹ�� CDATASection 
    '���ֻʹ�� Response.write ���� XML �ļ�����, ��Ų���Ҫʹ�� CDATASECTION, Response.write �Զ� ת�� 
    '����� ASCII �ַ�Ϊ HTML �ַ� 
    '���ʹ�� FSO ��ȡ XML ����, �е�ţͷ���������ζ�� 
%> 
