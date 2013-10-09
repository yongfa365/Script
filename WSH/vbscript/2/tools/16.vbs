'***********************************************
'*by r05e
'*bios信息
'***********************************************
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
   
    doc1.writeln "<html><head><title>显示系统BIOS信息</title></head>"
    doc1.writeln "<body bgcolor='silver'><pre><center><font color=red size=7>BIOS信息</font></center><p><font color=blue size=3>"
   
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")
For each objBIOS in colBIOS
    doc1.writeln "Build Number: " &objBIOS.BuildNumber
    doc1.writeln "Current Language: " &objBIOS.CurrentLanguage
    doc1.writeln "Installable Languages: " &objBIOS.InstallableLanguages
    doc1.writeln "Manufacturer: " &objBIOS.Manufacturer
    doc1.writeln "Name: "&objBIOS.Name
    doc1.writeln "Primary BIOS: "& objBIOS.PrimaryBIOS
    doc1.writeln "Release Date: " &objBIOS.ReleaseDate
    doc1.writeln "Serial Number: " & objBIOS.SerialNumber
    doc1.writeln "SMBIOS Version: " &objBIOS.SMBIOSBIOSVersion
    
    doc1.writeln "SMBIOS Minor Version: " &objBIOS.SMBIOSMinorVersion
    doc1.writeln "SMBIOS Present: " &objBIOS.SMBIOSPresent
    doc1.writeln "Status: " &objBIOS.Status
    doc1.writeln "Version: " &objBIOS.Version
    doc1.writeln "BIOS Characteristics: "
    For i = 0 to Ubound(objBIOS.BiosCharacteristics)
        doc1.writeln objBIOS.BiosCharacteristics(i)
    Next
Next


    
doc1.writeln "</font></p></pre></body></html>"
    doc1.close                  

set oIE=nothing