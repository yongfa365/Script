Dim dtDate
dtDate = InputBox("��������Ҫ��������" ,"���ڼ�����" ,Date())
If dtDate <> "" Then
	If Weekday(dtDate) = 4 Then
		MsgBox "���ϰ�" ,vbinformation ,"���ڼ�����"
	Else
		MsgBox "�ϰ�" ,vbinformation ,"���ڼ�����"
	End If
End If