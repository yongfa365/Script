'*********************************************
'*by r05e 
'*用IE显示系统环境变量,视觉效果好一点
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
   
    doc1.writeln "<html><head><title>显示系统环境变量</title></head>"
    doc1.writeln "<body bgcolor='silver'><pre><center><font color=red size=5>系统环境变量</font></center><p><font color=blue size=3>"
   
    

Set wshshell = CreateObject("WScript.Shell")
For Each EnvirSYSTEM In wshshell.Environment("SYSTEM")
 enOutSYSTEM=enOutSYSTEM&"当前"&EnvirSYSTEM&vbCrlf
Next

For Each EnvirPROCESS In wshshell.Environment("PROCESS")
 enOutPROCESS=enOutPROCESS&"当前"&EnvirPROCESS&vbCrlf
Next

For Each EnvirUSER In wshshell.Environment("USER")
 enOutUSER=enOutUSER&"当前"&EnvirUSER&vbCrlf
Next

For Each EnvirVOLATILE In wshshell.Environment("VOLATILE")
 enOutVOLATILE=enOutVOLATILE&"当前"&EnvirVOLATILE&vbCrlf
Next
 doc1.writeln enOutSYSTEM&enOutPROCESS&enOutUSER&enOutVOLATILE
doc1.writeln "</font></p></pre></body></html>"
    doc1.close                  
set wshshell=nothing
set oIE=nothing