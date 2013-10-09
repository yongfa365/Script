Set HttpObj = CreateObject("AspHTTP.Conn")
HTTPObj.Url = "http://www.im286.com/index.php"
'HTTPObj.PostData = "suid=jimb&pwd=macabre&id=32&val=1.5"
HTTPObj.TimeOut = 1800
HTTPObj.Accept = "*/*"
HTTPObj.FollowRedirects = true
HTTPObj.Port = 80        
'HTTPObj.Proxy = "xxx.net:8080"                '使用代理地址，端口
'HTTPObj.ProxyPassword = "proxyusername:proxypassword"        '代理的用户名，密码
HttpObj.SaveFileTo = "D:/WEB/weburl/index.html"                '将远程页面保存到本地
HTTPObj.UserAgent = "Mozilla Compatible (MS IE 3.01 WinNT)"
HTTPObj.Protocol = "HTTP/1.1"
HTTPObj.Authorization = "USER:pass"        
HTTPObj.ContentType = "application/x-www-form-urlencoded"
HTTPObj.RequestMethod = "POST"
'HTTPObj.GetHREFs
'HTTPObj.RequestMethod = "HEAD"
strResult = HTTPObj.GetURL
msgbox strResult
