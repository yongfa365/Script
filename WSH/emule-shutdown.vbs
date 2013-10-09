dir="E:\Temp"
Set fso=CreateObject("Scripting.FileSystemObject")
cmd="shutdown -s -f -t 60"
Set ws=WScript.CreateObject("WScript.Shell")
Wscript.Echo "emule自动关机脚本已运行……"
count=0
do until count<-1
chksize = fso.GetFolder(dir).Size
If chksize=0 Then
ws.run cmd,0
End If
WScript.Sleep 120000
loop