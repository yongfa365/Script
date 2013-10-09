'/*=========================================================================
' * Intro       [简单]VBS实现添加IE工具栏按钮
' * FileName    AddIEBar.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/AddIEBar.vbs.html
' * MadeTime    2008-01-01 11:33:27
' * LastModify  2008-01-01 11:33:27
' *==========================================================================*/

On Error Resume Next

WScript.Echo "点确定开始设置，如果有下载图标的功能速度会有些慢"

'操作之前删除所有非系统自带的工具栏按钮。
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.regDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\"'

Save2Local "http://www.yongfa365.com/favicon.ico", "C:\WINDOWS\www.yongfa365.com.ico"'下载一个ico图片，如果本地有，可以不用这句
ButtonRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\{ABCDEfef-1234-2134-7980-213421423401}\"'注册表位置，位数一定要对，字符随便写，只要不和别的重名就可以了
ButtonText = "柳永法(yongfa365)'Blog" '按钮名称
DefaultShow = "yes" '按钮显示否
URL = "http://www.yongfa365.com/" '点击后打开地址，可以是本地，也可以是网址
Icon = "C:\WINDOWS\www.yongfa365.com.ico" '默认显示图标
HotIcon = "C:\WINDOWS\www.yongfa365.com.ico" '鼠标放上去显示图标
MenuText = "柳永法的精品Blog" '工具菜单下显示文字
MenuStatusBar = "柳永法的精品Blog，这个网站非常不错的，http://www.yongfa365.com/" '鼠标放在工具菜单相应菜单上时状态栏显示文字
Call AddIEbar(ButtonRegPath, ButtonText, DefaultShow, URL, Icon, HotIcon, MenuText, MenuStatusBar)

'Save2Local "http://www.yongfa365.com/Images/ICON/RAR.ico", "C:\WINDOWS\WinRAR.ico"'下载一个ico图片，如果本地有，可以不用这句
'ButtonRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\{ABCDEfef-1234-2134-7980-213421423402}\"'注册表位置，位数一定要对，字符随便写，只要不和别的重名就可以了
'ButtonText = "打开WinRAR" '按钮名称
'DefaultShow = "yes" '按钮显示否
'URL = "C:\Program Files\WinRAR\WinRAR.exe" '点击后打开地址，可以是本地，也可以是网址
'Icon = "C:\WINDOWS\WinRAR.ico" '默认显示图标
'HotIcon = "C:\WINDOWS\WinRAR.ico" '鼠标放上去显示图标
'MenuText = "打开WinRAR" '工具菜单下显示文字
'MenuStatusBar = "打开WinRAR" '鼠标放在工具菜单相应菜单上时状态栏显示文字
'Call AddIEbar(ButtonRegPath, ButtonText, DefaultShow, URL, Icon, HotIcon, MenuText, MenuStatusBar)



'默认情况下即便是设置了显示，但系统还不会显示在工具栏上，得手动添加上去，下边两句实现自动添加（非完美解决方法）
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.regDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar\{1E796980-9CC5-11D1-A83F-00C04FC99D61}"

WScript.Echo "OK,操作完成"

'---------------------------------------------------添加到IE工具栏函数-----------------------------------

Function AddIEbar(ButtonRegPath, ButtonText, DefaultShow, URL, Icon, HotIcon, MenuText, MenuStatusBar)
    Set WSHShell = WScript.CreateObject("WScript.Shell")
    WSHShell.RegWrite ButtonRegPath&"CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}", "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"ButtonText", ButtonText, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"Exec", URL, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"Default Visible", DefaultShow, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"HotIcon", HotIcon, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"Icon", Icon, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"MenuText", MenuText, "REG_SZ"
    WSHShell.RegWrite ButtonRegPath&"MenuStatusBar", MenuStatusBar, "REG_SZ"
End Function

'---------------------------------------------------取图标操作函数---------------------------------------

Function getHTTPimg(url)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send()
    If xmlhttp.Status<>200 Then Exit Function
    getHTTPimg = xmlhttp.responseBody
    Set xmlhttp = Nothing
    If Err.Number<>0 Then Err.Clear
End Function

Function Save2Local(from, tofile)
    Dim geturl, objStream, imgs
    geturl = Trim(from)
    imgs = gethttpimg(geturl)
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objstream.Write imgs
    objstream.SaveToFile tofile, 2
    objstream.Close()
    Set objstream = Nothing
End Function
