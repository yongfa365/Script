Function WriteToFile (FileUrl, Str, CharSet)
    Set stm = CreateObject("Adodb.Stream")
    stm.Type = 2
    stm.mode = 3
    stm.charset = CharSet
    stm.Open
    stm.WriteText Str
    stm.SaveToFile FileUrl, 2
    stm.flush
    stm.Close
    Set stm = Nothing
End Function

Function DeleteFile(FileName)
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FSO.DeleteFile FileName, True
    End If
    If Err>0 Then
        Err.Clear
        DeleteFile = False
    Else
        DeleteFile = True
    End If
End Function

Function FileExits(FileName)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FileName) Then
        FileExits = True
    Else
        FileExits = False
    End If
End Function


EXT=INPUTBOX("任意您想要的扩展名，不能为空","请输入文件扩展名","vbs")
CHARSET=INPUTBOX("编码,如gb2312,utf-8","请输入文件编码","gb2312")
PRE=INPUTBOX("文件名前缀，为空表示只是数字","请输入文件名前缀","")
NO=INPUTBOX("您要生成多少个文件","多少个","4")

For i=1 to NO
if FileExits(PRE & i & "." & EXT)=False Then WriteToFile PRE & i & "." & EXT, "",CHARSET
'DeleteFile PRE & i & "." & EXT
next
