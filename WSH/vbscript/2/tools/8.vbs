'*********************************************
'*by r05e （从Adam的改编过来的，Adam的显示方式不好）
'*用消息框显示系统环境变量
'*********************************************
Set wshshell = CreateObject("WScript.Shell")
For Each EnvirSYSTEM In wshshell.Environment("SYSTEM")
 enOutSYSTEM=enOutSYSTEM&"当前"&EnvirSYSTEM&vbCrlf
Next

For Each EnvirPROCESS In wshshell.Environment("PROCESS")
 enOutPROCESS=enOutPROCESS&"当前"&EnvirPROCESS&vbCrlf
Next

For Each EnvirUSER In wshshell.Environment("USER")
 enOutUSER=enOutUSER&"当前"&EnvirUSER&vbCrlf
Next

For Each EnvirVOLATILE In wshshell.Environment("VOLATILE")
 enOutVOLATILE=enOutVOLATILE&"当前"&EnvirVOLATILE&vbCrlf
Next
WScript.Echo  enOutSYSTEM&enOutPROCESS&enOutUSER&enOutVOLATILE
set wshshell=nothing