'���ƣ�iis��־����������ط��ʴ���ͳ��vbs��
'���ߣ�������
'��ַ��www.yongfa365.com
'���ڣ�2007-6-18
'ʹ�÷�����������ļ���Ϊ��iislogAnalyzer.vbs,Ȼ�����б��

if WScript.Arguments.length<1 then 
	File =InputBox("������־Ŀ¼")
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
Call MsgBox(x,,title&"ͳ�ƽ��")

End Function