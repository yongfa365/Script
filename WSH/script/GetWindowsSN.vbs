'   ##############################################################
'   #         #
'   # VBScript to find the DigitalProductID for your   #
'   # Microsoft windows Installation and decode it to   #
'   # retrieve your windows Product Key     #
'   #         #
'   # -----------------------------------------------   #
'   #         #
'   #   Created by:   Parabellum    #
'   #         #
'   ##############################################################
'
' <--------------- Open Registry Key and populate binary data into an array -------------------------->
'
Const HKEY_LOCAL_MACHINE = &H80000002
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "DigitalProductId"
strComputer = "."
Dim iValues()
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
oReg.GetBinaryValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, iValues
Dim arrDPID
arrDPID = Array()
For i = 52 To 66
    ReDim Preserve arrDPID( UBound(arrDPID) + 1 )
    arrDPID( UBound(arrDPID) ) = iValues(i)
Next
' <--------------- Create an array to hold the valid characters for a microsoft Product Key -------------------------->
Dim arrChars
arrChars = Array("B", "C", "D", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "6", "7", "8", "9")

' <--------------- The clever bit !!! (Decrypt the base24 encoded binary data)-------------------------->
For i = 24 To 0 Step -1
    k = 0
    For j = 14 To 0 Step -1
        k = k * 256 Xor arrDPID(j)
        arrDPID(j) = Int(k / 24)
        k = k Mod 24
    Next
    strProductKey = arrChars(k) & strProductKey
    ' <------- add the "-" between the groups of 5 Char -------->
    If i Mod 5 = 0 And i <> 0 Then strProductKey = "-" & strProductKey
Next
strFinalKey = strProductKey
'
' <---------- This part of the script displays operating system Information and the license Key --------->
'
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    strOS = objOperatingSystem.Caption
    strBuild = objOperatingSystem.BuildNumber
    strSerial = objOperatingSystem.SerialNumber
    strRegistered = objOperatingSystem.RegisteredUser
Next
Set wshShell = CreateObject("wscript.shell")
strPopupMsg = strOS & vbNewLine & vbNewLine
strPopupMsg = strPopupMsg & "Build Number:   " & strBuild & vbNewLine
strPopupMsg = strPopupMsg & "PID:   " & strSerial & vbNewLine & vbNewLine
strPopupMsg = strPopupMsg & "Registered to:   " & strRegistered & vbNewLine & vbNewLine & vbNewLine
strPopupMsg = strPopupMsg & "Your Windows Product Key is:" & vbNewLine & vbNewLine & strFinalKey
strPopupTitle = "Microsoft Windows License Information"
wshShell.Popup strPopupMsg, , strPopupTitle, vbCancelOnly + vbInformation
WScript.Quit
