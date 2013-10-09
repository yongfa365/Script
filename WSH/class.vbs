
Class yongfa365
Public  Function CreateFile(FileName, Content)

Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fd = FSO.CreateTextFile(FileName, True)
    fd.WriteLine Content
    If Err>0 Then
        Err.Clear
        CreateFile = False
    Else
        CreateFile = True
    End If
End Function

end class

set a=new yongfa365
msgbox(a.CreateFile("123.txt","我的第一个类"))