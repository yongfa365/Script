FileName=wscript.arguments(0)

With CreateObject("adodb.stream")
	.Type=1
	.Open
	.loadfromfile FileName
str=.read
End With

With CreateObject("scripting.filesystemobject").opentextfile(FileName&".bat",2,True)
	For i=1 To lenb(str)
		bt=ascb(midb(str,i,1))
		If bt<16 Then .Write "0"
		.Write Hex(bt)
	Next
End With