'把这段代码复制到记事本里然后保存为：win2003-200K.vbs，看好了扩展名为.vbs   
Set providerObj=GetObject("winmgmts:/root/MicrosoftIISv2")   
Set vdirObj=providerObj.Get("IIsWebServiceSetting='W3SVC'")   
WScript.Echo "Before: " & vdirObj.AspMaxRequestEntityAllowed   
vdirObj.AspMaxRequestEntityAllowed=20480000 '可接收多大字节,此处默认为：204800即：200K   
vdirObj.Put_()   
WScript.Echo "Now: " & vdirObj.AspMaxRequestEntityAllowed 