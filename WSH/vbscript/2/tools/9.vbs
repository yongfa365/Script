'*********************************************
'*by r05e 
'*��IE��ʾϵͳ��������,�Ӿ�Ч����һ��
'*********************************************

Dim oIE, doc1
    
    Set oIE = WScript.CreateObject("InternetExplorer.Application")
    oIE.Navigate "about:blank"  
    oIE.Visible = 1             
    oIE.ToolBar = 0
    oIE.StatusBar = 0
    oIE.Width=750
    oIE.Height=700
    Do While (oIE.Busy): Loop   
 
    Set doc1 = oIE.Document     
    doc1.open                   
   
    doc1.writeln "<html><head><title>��ʾϵͳ��������</title></head>"
    doc1.writeln "<body bgcolor='silver'><pre><center><font color=red size=5>ϵͳ��������</font></center><p><font color=blue size=3>"
   
    

Set wshshell = CreateObject("WScript.Shell")
For Each EnvirSYSTEM In wshshell.Environment("SYSTEM")
 enOutSYSTEM=enOutSYSTEM&"��ǰ"&EnvirSYSTEM&vbCrlf
Next

For Each EnvirPROCESS In wshshell.Environment("PROCESS")
 enOutPROCESS=enOutPROCESS&"��ǰ"&EnvirPROCESS&vbCrlf
Next

For Each EnvirUSER In wshshell.Environment("USER")
 enOutUSER=enOutUSER&"��ǰ"&EnvirUSER&vbCrlf
Next

For Each EnvirVOLATILE In wshshell.Environment("VOLATILE")
 enOutVOLATILE=enOutVOLATILE&"��ǰ"&EnvirVOLATILE&vbCrlf
Next
 doc1.writeln enOutSYSTEM&enOutPROCESS&enOutUSER&enOutVOLATILE
doc1.writeln "</font></p></pre></body></html>"
    doc1.close                  
set wshshell=nothing
set oIE=nothing