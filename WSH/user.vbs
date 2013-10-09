msgbox getContents("a(.+)b", "a23234b ab a67896896b sadfasdfb" ,True)(0)
Function getContents(patrn, strng ,yinyong)
    Set re = New RegExp
    re.Pattern = patrn
    re.IgnoreCase = True
    re.Global = True
    Set Matches = re.Execute(strng)
    if yinyong then
	    For i=0 to Matches.count-1
	    	if Matches(i).value<>"" then RetStr = RetStr & Matches(i).SubMatches(0) & "ÁøÓÀ·¨"
	    Next
	Else
	    For each oMatch in Matches
	    	if oMatch.value<>"" then RetStr = RetStr & oMatch.value & "ÁøÓÀ·¨"
	    Next
	end if
    getContents = split(RetStr,"ÁøÓÀ·¨")
End Function