'/*=========================================================================
' * Intro       [��]VBSʵ�����IE��������ť
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

WScript.Echo "��ȷ����ʼ���ã����������ͼ��Ĺ����ٶȻ���Щ��"

'����֮ǰɾ�����з�ϵͳ�Դ��Ĺ�������ť��
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.regDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\"'

Save2Local "http://www.yongfa365.com/favicon.ico", "C:\WINDOWS\www.yongfa365.com.ico"'����һ��icoͼƬ����������У����Բ������
ButtonRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\{ABCDEfef-1234-2134-7980-213421423401}\"'ע���λ�ã�λ��һ��Ҫ�ԣ��ַ����д��ֻҪ���ͱ�������Ϳ�����
ButtonText = "������(yongfa365)'Blog" '��ť����
DefaultShow = "yes" '��ť��ʾ��
URL = "http://www.yongfa365.com/" '�����򿪵�ַ�������Ǳ��أ�Ҳ��������ַ
Icon = "C:\WINDOWS\www.yongfa365.com.ico" 'Ĭ����ʾͼ��
HotIcon = "C:\WINDOWS\www.yongfa365.com.ico" '������ȥ��ʾͼ��
MenuText = "�������ľ�ƷBlog" '���߲˵�����ʾ����
MenuStatusBar = "�������ľ�ƷBlog�������վ�ǳ�����ģ�http://www.yongfa365.com/" '�����ڹ��߲˵���Ӧ�˵���ʱ״̬����ʾ����
Call AddIEbar(ButtonRegPath, ButtonText, DefaultShow, URL, Icon, HotIcon, MenuText, MenuStatusBar)

'Save2Local "http://www.yongfa365.com/Images/ICON/RAR.ico", "C:\WINDOWS\WinRAR.ico"'����һ��icoͼƬ����������У����Բ������
'ButtonRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\{ABCDEfef-1234-2134-7980-213421423402}\"'ע���λ�ã�λ��һ��Ҫ�ԣ��ַ����д��ֻҪ���ͱ�������Ϳ�����
'ButtonText = "��WinRAR" '��ť����
'DefaultShow = "yes" '��ť��ʾ��
'URL = "C:\Program Files\WinRAR\WinRAR.exe" '�����򿪵�ַ�������Ǳ��أ�Ҳ��������ַ
'Icon = "C:\WINDOWS\WinRAR.ico" 'Ĭ����ʾͼ��
'HotIcon = "C:\WINDOWS\WinRAR.ico" '������ȥ��ʾͼ��
'MenuText = "��WinRAR" '���߲˵�����ʾ����
'MenuStatusBar = "��WinRAR" '�����ڹ��߲˵���Ӧ�˵���ʱ״̬����ʾ����
'Call AddIEbar(ButtonRegPath, ButtonText, DefaultShow, URL, Icon, HotIcon, MenuText, MenuStatusBar)



'Ĭ������¼�������������ʾ����ϵͳ��������ʾ�ڹ������ϣ����ֶ������ȥ���±�����ʵ���Զ���ӣ����������������
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.regDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar\{1E796980-9CC5-11D1-A83F-00C04FC99D61}"

WScript.Echo "OK,�������"

'---------------------------------------------------��ӵ�IE����������-----------------------------------

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

'---------------------------------------------------ȡͼ���������---------------------------------------

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
