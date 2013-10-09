n = InputBox("请输入一个正数", "输入", 3389)

MSG = MSG & vbcrlf & n & "的 2进制形式是:" & change (n, 2)
MSG = MSG & vbcrlf & n & "的 8进制形式是:" & change (n, 8)
MSG = MSG & vbcrlf & n & "的16进制形式是:" & change (n, 16)
Msgbox(MSG)

Function change(byval n, byval sysn)
    While (n)
        t = n Mod sysn
        change = Mid("0123456789ABCDEF", t + 1, 1) & change
        n = Fix(n / sysn)
    Wend
End Function