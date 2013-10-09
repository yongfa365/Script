'********************************
'*by r05e
'*启动脚本（信息提示框）
'********************************
Function loginmsg()
msg=InputBox("请输入要输入的消息内容，它将在每个用户LOGIN时出现","输入消息", "注意不要访问非法网络资源", 100, 100)
title=InputBox("请输入消息的标题","输入标题", "管理员提示", 100, 100)
prompt=InputBox("请输入系统盘位置","输入系统盘位置", "C", 100, 100)
If msg<>"" and title<>"" and prompt<>"" Then
If Len(prompt)=1 and Asc(Lcase(prompt))>96 and Asc(Lcase(prompt))<123 Then
set fso=createobject("scripting.FileSystemObject")
set crVBS=fso.createtextfile(prompt&":\Documents and Settings\All Users\「开始」菜单\程序\启动\msg.vbs",True)

crVBS.writeline("MsgBox "&Chr(34)&msg&Chr(34)&",64,"&Chr(34)&title&Chr(34))
crVBS.close
wscript.echo "操作成功"
End If
ElseIF msg="" and title="" and prompt="" Then
wscript.echo "操作已经取消"
Else wscript.echo "输入错误，请重试"
loginmsg()
End If
set fso=nothing
set crVBS=nothing
End Function
loginmsg()