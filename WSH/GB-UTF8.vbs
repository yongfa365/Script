
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

'∫∫◊÷◊™ªªŒ™UTF-8

Function chinese2unicode(Str)
    Dim i
    Dim Str_one
    Dim Str_unicode
    For i = 1 To Len(Str)
        Str_one = Mid(Str, i, 1)
        Str_unicode = Str_unicode&Chr(38)
        Str_unicode = Str_unicode&Chr(35)
        Str_unicode = Str_unicode&Chr(120)
        Str_unicode = Str_unicode& Hex(ascw(Str_one))
        Str_unicode = Str_unicode&Chr(59)
    Next
    chinese2unicode = Str_unicode
End Function

'UTF-8 To GB2312

Function UTF2GB(UTFStr)
    For Dig = 1 To Len(UTFStr)
        If Mid(UTFStr, Dig, 1) = "%" Then
            If Len(UTFStr) >= Dig + 8 Then
                GBStr = GBStr & ConvChinese(Mid(UTFStr, Dig, 9))
                Dig = Dig + 8
            Else
                GBStr = GBStr & Mid(UTFStr, Dig, 1)
            End If
        Else
            GBStr = GBStr & Mid(UTFStr, Dig, 1)
        End If
    Next
    UTF2GB = GBStr
End Function


Function ConvChinese(x)
    A = Split(Mid(x, 2), "%")
    i = 0
    j = 0
    
    For i = 0 To UBound(A)
        A(i) = c16to2(A(i))
    Next
    
    For i = 0 To UBound(A) -1
        DigS = InStr(A(i), "0")
        Unicode = ""
        For j = 1 To DigS -1
            If j = 1 Then
                A(i) = Right(A(i), Len(A(i)) - DigS)
                Unicode = Unicode & A(i)
            Else
                i = i + 1
                A(i) = Right(A(i), Len(A(i)) -2)
                Unicode = Unicode & A(i)
            End If
        Next
        
        If Len(c2to16(Unicode)) = 4 Then
            ConvChinese = ConvChinese & chrw(Int("&H" & c2to16(Unicode)))
        Else
            ConvChinese = ConvChinese & Chr(Int("&H" & c2to16(Unicode)))
        End If
    Next
End Function


Function c2to16(x)
    i = 1
    For i = 1 To Len(x) step 4
        c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
    Next
End Function

Function c2to10(x)
    c2to10 = 0
    If x = "0" Then Exit Function
    i = 0
    For i = 0 To Len(x) -1
        If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2^(i)
    Next
End Function


Function c16to2(x)
    i = 0
    For i = 1 To Len(Trim(x))
        tempstr = c10to2(CInt(Int("&h" & Mid(x, i, 1))))
        Do While Len(tempstr)<4
            tempstr = "0" & tempstr
        Loop
        c16to2 = c16to2 & tempstr
    Next
End Function


Function c10to2(x)
    mysign = Sgn(x)
    x = Abs(x)
    DigS = 1
    Do
        If x<2^DigS Then
            Exit Do
        Else
            DigS = DigS + 1
        End If
    Loop
    tempnum = x
    
    i = 0
    For i = DigS To 1 step -1
        If tempnum>= 2^(i -1) Then
            tempnum = tempnum -2^(i -1)
            c10to2 = c10to2 & "1"
        Else
            c10to2 = c10to2 & "0"
        End If
    Next
    If mysign = -1 Then c10to2 = "-" & c10to2
End Function

MsgBox chinese2unicode("Œ“")
