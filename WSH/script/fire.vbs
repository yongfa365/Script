'******************************************************************************
'WinFire-com.vbs
'Author: Peter Costantini, The Microsoft Scripting Guys
'Date: 8/26/04
'Version: 1.0
'This script opens specified ports and authorizes specified applications
'in Windows Firewall on the local computer.
'******************************************************************************

'Set constants.
Const NET_FW_IP_PROTOCOL_TCP = 6
Const NET_FW_IP_PROTOCOL_UDP = 17
Const NET_FW_SCOPE_ALL = 0
Const NET_FW_SCOPE_LOCAL_SUBNET = 1
'First dimension of arrNewPorts must equal # of ports to be added minus 1.
Dim arrNewPorts(2,4)
'First dimension of arrNewApps must equal # of apps to be allowed.
Dim arrNewApps(2,3)

'Edit this list to add ports to the exceptions list.
'Scope and Enabled are optional properties.

arrNewPorts(0,0) = "FPS" 'Name
arrNewPorts(0,1) = 137 'Port
arrNewPorts(0,2) = NET_FW_IP_PROTOCOL_TCP 'Protocol
arrNewPorts(0,3) = NET_FW_SCOPE_ALL 'Scope - default is NET_FW_SCOPE_ALL
arrNewPorts(0,4) = True 'Enabled - default is True

arrNewPorts(1,0) = "FPS1"
arrNewPorts(1,1) = 138
arrNewPorts(1,2) = NET_FW_IP_PROTOCOL_UDP
arrNewPorts(1,3) = NET_FW_SCOPE_ALL
arrNewPorts(1,4) = True

arrNewPorts(2,0) = "XXX"
arrNewPorts(2,1) = 552
arrNewPorts(2,2) = NET_FW_IP_PROTOCOL_UDP
arrNewPorts(2,3) = NET_FW_SCOPE_LOCAL_SUBNET
arrNewPorts(2,4) = True

'Edit this list to add programs to the exceptions list.
'Scope and Enabled are optional properties.

arrNewApps(0,0) = "NsLookup" 'Name
arrNewApps(0,1) = "%windir%\system32\nslookup.exe" 'ProcessImageFileName
'Must be a fully qualified path, but can contain environment variables.
arrNewApps(0,2) = NET_FW_SCOPE_ALL 'Scope - default is NET_FW_SCOPE_ALL
arrNewApps(0,3) = True 'Enabled

arrNewApps(1,0) = "Notepad"
arrNewApps(1,1) = "%windir%\system32\notepad.exe"
arrNewApps(1,2) = NET_FW_SCOPE_LOCAL_SUBNET
arrNewApps(1,3) = True

arrNewApps(2,0) = "Calculator"
arrNewApps(2,1) = "%windir%\system32\calc.exe"
arrNewApps(2,2) = NET_FW_SCOPE_ALL
arrNewApps(2,3) = True

On Error Resume Next
'Create the firewall manager object.
Set objFwMgr = CreateObject("HNetCfg.FwMgr")
If Err <> 0 Then
  WScript.Echo "Unable to connect to Windows Firewall."
  WScript.Quit
End If
'Get the current profile for the local firewall policy.
Set objProfile = objFwMgr.LocalPolicy.CurrentProfile
Set colOpenPorts = objProfile.GloballyOpenPorts
Set colAuthorizedApps = objProfile.AuthorizedApplications

WScript.Echo VbCrLf & "New open ports added:"
For i = 0 To UBound(arrNewPorts)
'Create an FWOpenPort object
  Set objOpenPort = CreateObject("HNetCfg.FWOpenPort")
  objOpenPort.Name = arrNewPorts(i, 0)
  objOpenPort.Port = arrNewPorts(i, 1)
  objOpenPort.Protocol = arrNewPorts(i, 2)
  objOpenPort.Scope = arrNewPorts(i, 3)
  objOpenPort.Enabled = arrNewPorts(i, 4)
  colOpenPorts.Add objOpenPort
  If Err = 0 Then
    WScript.Echo VbCrLf & "Name: " & objOpenPort.Name
    WScript.Echo "  Protocol: " & objOpenPort.Protocol
    WScript.Echo "  Port Number: " & objOpenPort.Port
    WScript.Echo "  Scope: " & objOpenPort.Scope
    WScript.Echo "  Enabled: " & objOpenPort.Enabled
  Else
    WScript.Echo VbCrLf & "Unable to add port: " & arrNewPorts(i, 0)
    WScript.Echo "  Error Number:" & Err.Number
    WScript.Echo "  Source:" & Err.Source
    WScript.Echo "  Description:" & Err.Description
  End If
  Err.Clear
Next

WScript.Echo VbCrLf & "New authorized applications added:"
For i = 0 To UBound(arrNewApps)
'Create an FwAuthorizedApplication object
  Set objAuthorizedApp = CreateObject("HNetCfg.FwAuthorizedApplication")
  objAuthorizedApp.Name = arrNewApps(i,0)
  objAuthorizedApp.ProcessImageFileName = arrNewApps(i, 1)
  objAuthorizedApp.Scope = arrNewApps(i, 2)
  objAuthorizedApp.Enabled = arrNewApps(i, 3)
  colAuthorizedApps.Add objAuthorizedApp
  If Err = 0 Then
    WScript.Echo VbCrLf & "Name: " & objAuthorizedApp.Name
    WScript.Echo "  Process Image File: " & _
     objAuthorizedApp.ProcessImageFileName
    WScript.Echo "  Scope: " & objAuthorizedApp.Scope
    WScript.Echo "  Enabled: " & objAuthorizedApp.Enabled
  Else
    WScript.Echo VbCrLf & "Unable to add application: " & arrNewApps(i,0)
    WScript.Echo "  Error Number:" & Err.Number
    WScript.Echo "  Source:" & Err.Source
    WScript.Echo "  Description:" & Err.Description
  End If
Next

Set colOpenPorts = objProfile.GloballyOpenPorts
WScript.Echo VbCrLf & "All listed ports after operation:"
For Each objPort In colOpenPorts
  WScript.Echo VbCrLf & "Name: " & objPort.Name
  WScript.Echo "  Protocol: " & objPort.Protocol
  WScript.Echo "  Port Number: " & objPort.Port
  WScript.Echo "  Scope: " & objPort.Scope
  WScript.Echo "  Enabled: " & objPort.Enabled
Next

Set colAuthorizedApps = objProfile.AuthorizedApplications
WScript.Echo VbCrLf & "All listed applications after operation:"
For Each objApp In colAuthorizedApps
  WScript.Echo VbCrLf & "Name: " & objApp.Name
  WScript.Echo "  Process Image File: " & objApp.ProcessImageFileName
  WScript.Echo "  Scope: " & objApp.Scope
  WScript.Echo "  Enabled: " & objApp.Enabled
Next
