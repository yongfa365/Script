'����δ��븴�Ƶ����±���Ȼ�󱣴�Ϊ��win2003-200K.vbs����������չ��Ϊ.vbs   
Set providerObj=GetObject("winmgmts:/root/MicrosoftIISv2")   
Set vdirObj=providerObj.Get("IIsWebServiceSetting='W3SVC'")   
WScript.Echo "Before: " & vdirObj.AspMaxRequestEntityAllowed   
vdirObj.AspMaxRequestEntityAllowed=20480000 '�ɽ��ն���ֽ�,�˴�Ĭ��Ϊ��204800����200K   
vdirObj.Put_()   
WScript.Echo "Now: " & vdirObj.AspMaxRequestEntityAllowed 