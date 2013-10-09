Function getContent(patrn, strng)
    Set re= New RegExp '建立正则表达式。
    re.Pattern = patrn'设置模式。
    re.IgnoreCase = True '设置是否区分字符大小写。
    re.Global = True '设置全局可用性。
    Set Matches = re.Execute(strng)'执行搜索。
    RetStr0= Matches.Count
    For Each oMatch in Matches'遍历匹配集合。
        RetStr=RetStr & oMatch.value
        RetStr1=RetStr1 & oMatch.SubMatches(0)
        RetStr2=RetStr2 & oMatch.length
    Next
    getContent ="Matches.Count:" & RetStr0 & vbnewline & "oMatch.value:" & RetStr & vbnewline & "oMatch.SubMatches(0):" & RetStr1 & vbnewline  & "oMatch.length:" & RetStr2
End Function

msgbox getContent("<title>(.+?)</title>", "<title>1</title><title>22</title><title>333</title><title>4444</title><title>55555</title>")