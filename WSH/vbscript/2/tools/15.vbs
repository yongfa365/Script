'***********************************************
'*bY r05e
'*显示HOT-FIXES
'***********************************************

Dim oIE, doc1
    
    Set oIE = WScript.CreateObject("InternetExplorer.Application")
    oIE.Navigate "about:blank"  
    oIE.Visible = 1             
    oIE.ToolBar = 0
    oIE.StatusBar = 0
    oIE.Width=750
    oIE.Height=500
    Do While (oIE.Busy): Loop   
 
    Set doc1 = oIE.Document     
    doc1.open                   
   
    doc1.writeln "<html><head><title>显示系统的Hot-Fixes</title></head>"
    doc1.writeln "<body bgcolor='silver'><pre><center><font color=red size=15>HOT-FIXES</font></center><p><font color=blue size=3>"
   

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")
For Each objQuickFix in colQuickFixes
    doc1.writeln "Computer: " & objQuickFix.CSName
    doc1.writeln "Description: " & objQuickFix.Description
    doc1.writeln "Hot Fix ID: " & objQuickFix.HotFixID
    doc1.writeln "Installation Date: " & objQuickFix.InstallDate
    doc1.writeln "Installed By: " & objQuickFix.InstalledBy
Next

doc1.writeln "</font></p></pre></body></html>"
    doc1.close                  

set oIE=nothing
