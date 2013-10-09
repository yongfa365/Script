Dim dtDate
dtDate = InputBox("请输入需要检查的日期" ,"日期检测程序" ,Date())
If dtDate <> "" Then
	If Weekday(dtDate) = 4 Then
		MsgBox "不上班" ,vbinformation ,"日期检测程序"
	Else
		MsgBox "上班" ,vbinformation ,"日期检测程序"
	End If
End If