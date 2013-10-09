on error resume next
Function DateFormat(str)
    Set re = New RegExp '建立正则表达式。
    re.Pattern = "(\d{4})[-/年]*(\d{2})[-/月]*(\d{2})[-/日]*.*(\d{2}):(\d{2}).*"'设置模式。
    re.IgnoreCase = True '设置是否区分字符大小写。
    re.Global = True '设置全局可用性。
    Set Matches = re.Execute(str)'执行搜索。
    
	with Matches(0)
		a0=.SubMatches(0)
		a1=.SubMatches(1)
		a2=.SubMatches(2)
		a3=.SubMatches(3)
		a4=.SubMatches(4)

	end with
msgbox a0
msgbox a1
msgbox a2
msgbox a3
msgbox a4


end Function

DateFormat "2003-12-12 12:35"