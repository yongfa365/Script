Set HttpObj = CreateObject("AspHTTP.Conn")
HTTPObj.Url = "http://www.im286.com/index.php"
'HTTPObj.PostData = "suid=jimb&pwd=macabre&id=32&val=1.5"
HTTPObj.TimeOut = 1800
HTTPObj.Accept = "*/*"
HTTPObj.FollowRedirects = true
HTTPObj.Port = 80        
'HTTPObj.Proxy = "xxx.net:8080"                'ʹ�ô����ַ���˿�
'HTTPObj.ProxyPassword = "proxyusername:proxypassword"        '������û���������
HttpObj.SaveFileTo = "D:/WEB/weburl/index.html"                '��Զ��ҳ�汣�浽����
HTTPObj.UserAgent = "Mozilla Compatible (MS IE 3.01 WinNT)"
HTTPObj.Protocol = "HTTP/1.1"
HTTPObj.Authorization = "USER:pass"        
HTTPObj.ContentType = "application/x-www-form-urlencoded"
HTTPObj.RequestMethod = "POST"
'HTTPObj.GetHREFs
'HTTPObj.RequestMethod = "HEAD"
strResult = HTTPObj.GetURL
msgbox strResult
