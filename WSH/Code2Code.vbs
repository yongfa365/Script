n = InputBox("������һ������", "����", 3389)

MSG = MSG & vbcrlf & n & "�� 2������ʽ��:" & change (n, 2)
MSG = MSG & vbcrlf & n & "�� 8������ʽ��:" & change (n, 8)
MSG = MSG & vbcrlf & n & "��16������ʽ��:" & change (n, 16)
Msgbox(MSG)

Function change(byval n, byval sysn)
    While (n)
        t = n Mod sysn
        change = Mid("0123456789ABCDEF", t + 1, 1) & change
        n = Fix(n / sysn)
    Wend
End Function