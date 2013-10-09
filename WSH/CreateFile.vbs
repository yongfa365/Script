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


EXT=INPUTBOX("��������Ҫ����չ��������Ϊ��","�������ļ���չ��","vbs")
CHARSET=INPUTBOX("����,��gb2312,utf-8","�������ļ�����","gb2312")
PRE=INPUTBOX("�ļ���ǰ׺��Ϊ�ձ�ʾֻ������","�������ļ���ǰ׺","")
NO=INPUTBOX("��Ҫ���ɶ��ٸ��ļ�","���ٸ�","4")

For i=1 to NO
if FileExits(PRE & i & "." & EXT)=False Then WriteToFile PRE & i & "." & EXT, "",CHARSET
'DeleteFile PRE & i & "." & EXT
next
