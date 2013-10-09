'/*=========================================================================
' * Intro       网上找了一圈，都不怎么好，有一个比较不错的，汉化作者汉化时加了个自己的介绍文件，这个文件比程序本身还大，感觉不爽，于是本人的VBS版MAC修改代码便诞生了,在使用过程中如果出现不能上网的情况得返回一下网卡驱动（有些机器比较特别），如果要返回以前的MAC可以：开始-->控制面板-->网络连接-->点击您的网卡(一般是"本地连接")-->点击常规里的属性-->配置..-->高级-->选中-->NetworkAddress-->右边选择"不存在"
' * FileName    ChangeMAC.vbs
' * Author      yongfa365
' * Version     v3.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/ChangeMAC.vbs.html
' * MadeTime    2007-12-09 22:17:58
' * LastModify  2007-12-13 18:35:58
' *==========================================================================*/

On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=true", , 48)
For Each objItem in colItems
    msg = msg & "编号：" & objItem.Index & "   MAC：" & objItem.MACAddress & vbCrLf & "网卡：" & objItem.Description & vbCrLf & vbCrLf
Next

idx = InputBox( msg , "1/2请输入您要修改的MAC的编号", "1")
If Not IsNumeric(idx) Or Len(idx) = 0 Then
    WScript.Echo "编号输入有误，退出"
    Wscript.Quit
End If
MAC = InputBox( "输入你指定的MAC地址值（注意应该是12位的连续数字或字母，其间没有-、：等分隔符）" , "2/2请输入修改后的MAC地址", "000000000000")
MAC = Replace(Replace(Replace(MAC, ":", ""), "-", ""), " ", "")
If RegExpTest("[^\da-fA-F]", MAC)>0 Or Len(MAC)<>12 Then
    WScript.Echo "MAC输入有误，退出"
    Wscript.Quit
End If


idx = Right("00000"&idx, 4)
reg = "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\" & idx
Set WSHShell = CreateObject("WScript.Shell")
WshShell.RegWrite reg & "\NetworkAddress", MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\default" , MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\ParamDesc" , "NetworkAddress" , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\optional" , "1" , "REG_SZ"
'得到网卡的名称，比如“本地连接 2”
NetWorkName = WshShell.RegRead("HKLM\SYSTEM\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & WshShell.RegRead(reg & "\NetCfgInstanceId") & "\Connection\Name")

restartNetWork NetWorkName
'WScript.Echo "修改成功"

Function restartNetWork(sConnectionName)
    '重启网卡
    'sConnectionName = "本地连接 5" '可改成需要控制的连接名称，如"无线网络连接"等
    '定位到网络连接
    Set shellApp = CreateObject("shell.application")
    Set oControlPanel = shellApp.Namespace(3)
    For Each folderitem in oControlPanel.Items
        If RegExpTest("网络连接|network",folderitem.Name)<>0 Then
            Set oNetConnections = folderitem.GetFolder
            Exit For
        End If
    Next
    '定位到要处理的网卡
    For Each folderitem in oNetConnections.Items
        If LCase(folderitem.Name) = LCase(sConnectionName) Then
            Set oLanConnection = folderitem
            Exit For
        End If
    Next
    '重启网卡
    bEnable=True
    ExitOK=False
    Do while 1
        For Each verb in oLanConnection.verbs
	        if RegExpTest("禁用|disabled",verb.Name)<>0 and bEnable=True then
	                verb.DoIt
	                bEnable=False
	        elseif  RegExpTest("启用|enable",verb.Name)<>0 and bEnable=False then
	                verb.DoIt
	                ExitOK=True
	        end if
	        Exit For
        Next
        If ExitOK then exit do
     loop
WScript.Sleep 3000
CreateObject("wscript.shell").popup "成功",3
End Function


'正则测试有没有匹配内容

Function RegExpTest(patrn, strng)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    RegExpTest = Matches.Count
End Function