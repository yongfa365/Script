'名称：iis日志搜索引擎相关访问次数统计vbs版
'作者：柳永法
'网址：www.yongfa365.com
'日期：2007-6-18
'使用方法：把这个文件存为：iislogAnalyzer.vbs,然后运行便可

if WScript.Arguments.length<1 then 
	File =InputBox("输入日志目录")
	if File<>"" then checkSearchNO(File)
else
	for each File in WScript.Arguments
		checkSearchNO(File)
	next
end if


Function readFile(Filename)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set cnt = FSO.OpenTextFile(Filename, 1, True)
    readFile = cnt.ReadAll
End Function

Function RegExpTest(patrn, strng)
    Set regEx = New RegExp
    regEx.Pattern = patrn
    regEx.IgnoreCase = True
    regEx.Global = True
    Set Matches = regEx.Execute(strng)
    RegExpTest = Matches.count
End Function

Function checkSearchNO(File)
bots = "Baiduspider|Googlebot|yahoo|sogou|YodaoBot|msnbot|iaskspider"
bots = Split(bots, "|")
title_temp=split(File,"\")
title=title_temp(Ubound(title_temp)) & vbCrLf & vbCrLf
x=""
For i = LBound(bots) To UBound(bots)
    x = x & bots(i) & ":" & RegExpTest(bots(i),readFile(File)) & vbCrLf & vbCrLf
Next
Call MsgBox(x,,title&"统计结果")

End Function