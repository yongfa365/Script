'QuotedPrintable
Function QPEncoding(v, f)
    Dim s, t, i, j, h, l, x
    s = ""
    x = Len(v)
    For i = 1 To x
        t = Mid(v, i, 1)
        j = Asc(t)
        If j> 0 Then
            If f Then
                s = s & "=" & Right("00" & Hex(Asc(t)), 2)
            Else
                s = s & t
            End If
        Else
            If j < 0 Then j = j + &H10000
            h = (j And &HFF00) \ &HFF
            l = j And &HFF
            s = s & "=" & Hex(h) & "=" & Hex(l)
        End If
    Next
    QPEncoding = s
End Function


Function QPDecoding(Sin)
    Dim s, i, l, c, t, n
    s = ""
    l = Len(Sin)
    For i = 1 To l
        c = Mid(Sin, i, 1)
        If c<>"=" Then
            s = s & c
        Else
            c = Mid(Sin, i + 1, 2)
            i = i + 2
            t = CInt("&H" & c)
            If t<&H80 Then
                s = s & Chr(t)
            Else
                c = Mid(Sin, i + 1, 3)
                If Left(c, 1)<>"=" Then
                    QPDecoding = s
                    Exit Function
                Else
                    c = Right(c, 2)
                    n = CInt("&H" & c)
                    t = t * 256 + n -65536
                    s = s & Chr(t)
                    i = i + 3
                End If
            End If
        End If
    Next
    QPDecoding = s
End Function
