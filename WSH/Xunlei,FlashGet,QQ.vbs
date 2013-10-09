
sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
sBASE_64_CHARACTERS = strUnicode2Ansi(sBASE_64_CHARACTERS)

Function strUnicodeLen(asContents)
    '计算unicode字符串的Ansi编码的长度
    asContents1 = "a"&asContents
    len1 = Len(asContents1)
    k = 0
    For i = 1 To len1
        asc1 = Asc(Mid(asContents1, i, 1))
        If asc1<0 Then asc1 = 65536 + asc1
        If asc1>255 Then
            k = k + 2
        Else
            k = k + 1
        End If
    Next
    strUnicodeLen = k -1
End Function

Function strUnicode2Ansi(asContents)
    '将Unicode编码的字符串，转换成Ansi编码的字符串
    strUnicode2Ansi = ""
    len1 = Len(asContents)
    For i = 1 To len1
        varchar = Mid(asContents, i, 1)
        varasc = Asc(varchar)
        If varasc<0 Then varasc = varasc + 65536
        If varasc>255 Then
            varHex = Hex(varasc)
            varlow = Left(varHex, 2)
            varhigh = Right(varHex, 2)
            strUnicode2Ansi = strUnicode2Ansi & chrb("&H" & varlow ) & chrb("&H" & varhigh )
        Else
            strUnicode2Ansi = strUnicode2Ansi & chrb(varasc)
        End If
    Next
End Function

Function strAnsi2Unicode(asContents)
    '将Ansi编码的字符串，转换成Unicode编码的字符串
    strAnsi2Unicode = ""
    len1 = lenb(asContents)
    If len1 = 0 Then Exit Function
    For i = 1 To len1
        varchar = midb(asContents, i, 1)
        varasc = ascb(varchar)
        If varasc > 127 Then
            strAnsi2Unicode = strAnsi2Unicode & Chr(ascw(midb(asContents, i + 1, 1) & varchar))
            i = i + 1
        Else
            strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
        End If
    Next
End Function

Function Base64encode(asContents)
    '将Ansi编码的字符串进行Base64编码
    'asContents应当是ANSI编码的字符串（二进制的字符串也可以）
    Dim lnPosition
    Dim lsResult
    Dim Char1
    Dim Char2
    Dim Char3
    Dim Char4
    Dim Byte1
    Dim Byte2
    Dim Byte3
    Dim SaveBits1
    Dim SaveBits2
    Dim lsGroupBinary
    Dim lsGroup64
    Dim m4, len1, len2
    
    len1 = Lenb(asContents)
    If len1<1 Then
        Base64encode = ""
        Exit Function
    End If
    
    m3 = Len1 Mod 3
    If M3 > 0 Then asContents = asContents & String(3 - M3, chrb(0))
    '补足位数是为了便于计算
    
    If m3 > 0 Then
        len1 = len1 + (3 - m3)
        len2 = len1 -3
    Else
        len2 = len1
    End If
    
    lsResult = ""
    
    For lnPosition = 1 To len2 Step 3
        lsGroup64 = ""
        lsGroupBinary = Midb(asContents, lnPosition, 3)
        
        Byte1 = Ascb(Midb(lsGroupBinary, 1, 1))
        SaveBits1 = Byte1 And 3
        Byte2 = Ascb(Midb(lsGroupBinary, 2, 1))
        SaveBits2 = Byte2 And 15
        Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))
        
        Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) / 4) + 1, 1)
        Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) / 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) / 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
        Char4 = Midb(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
        lsGroup64 = Char1 & Char2 & Char3 & Char4
        
        lsResult = lsResult & lsGroup64
    Next
    
    '处理最后剩余的几个字符
    If M3 > 0 Then
        lsGroup64 = ""
        lsGroupBinary = Midb(asContents, len2 + 1, 3)
        
        Byte1 = Ascb(Midb(lsGroupBinary, 1, 1))
        SaveBits1 = Byte1 And 3
        Byte2 = Ascb(Midb(lsGroupBinary, 2, 1))
        SaveBits2 = Byte2 And 15
        Byte3 = Ascb(Midb(lsGroupBinary, 3, 1))
        
        Char1 = Midb(sBASE_64_CHARACTERS, ((Byte1 And 252) / 4) + 1, 1)
        Char2 = Midb(sBASE_64_CHARACTERS, (((Byte2 And 240) / 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
        Char3 = Midb(sBASE_64_CHARACTERS, (((Byte3 And 192) / 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
        
        If M3 = 1 Then
            lsGroup64 = Char1 & Char2 & ChrB(61) & ChrB(61) '用=号补足位数
        Else
            lsGroup64 = Char1 & Char2 & Char3 & ChrB(61) '用=号补足位数
        End If
        
        lsResult = lsResult & lsGroup64
    End If
    
    Base64encode = lsResult
    
End Function


Function Base64decode(asContents)
    '将Base64编码字符串转换成Ansi编码的字符串
    'asContents应当也是ANSI编码的字符串（二进制的字符串也可以）
    Dim lsResult
    Dim lnPosition
    Dim lsGroup64, lsGroupBinary
    Dim Char1, Char2, Char3, Char4
    Dim Byte1, Byte2, Byte3
    Dim M4, len1, len2
    
    len1 = Lenb(asContents)
    M4 = len1 Mod 4
    
    If len1 < 1 Or M4 > 0 Then
        '字符串长度应当是4的倍数
        Base64decode = ""
        Exit Function
    End If
    
    '判断最后一位是不是 = 号
    '判断倒数第二位是不是 = 号
    '这里m4表示最后剩余的需要单独处理的字符个数
    If midb(asContents, len1, 1) = chrb(61) Then m4 = 3
    If midb(asContents, len1 -1, 1) = chrb(61) Then m4 = 2
    
    If m4 = 0 Then
        len2 = len1
    Else
        len2 = len1 -4
    End If
    
    For lnPosition = 1 To Len2 Step 4
        lsGroupBinary = ""
        lsGroup64 = Midb(asContents, lnPosition, 4)
        Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1
        Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1
        Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1
        Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1
        Byte1 = Chrb(((Char2 And 48) / 16) Or (Char1 * 4) And &HFF)
        Byte2 = lsGroupBinary & Chrb(((Char3 And 60) / 4) Or (Char2 * 16) And &HFF)
        Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
        lsGroupBinary = Byte1 & Byte2 & Byte3
        
        lsResult = lsResult & lsGroupBinary
    Next
    
    '处理最后剩余的几个字符
    If M4 > 0 Then
        lsGroupBinary = ""
        lsGroup64 = Midb(asContents, len2 + 1, m4) & chrB(65) 'chr(65)=A，转换成值为0
        If M4 = 2 Then '补足4位，是为了便于计算
            lsGroup64 = lsGroup64 & chrB(65)
        End If
        Char1 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 1, 1)) - 1
        Char2 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 2, 1)) - 1
        Char3 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 3, 1)) - 1
        Char4 = InStrb(sBASE_64_CHARACTERS, Midb(lsGroup64, 4, 1)) - 1
        Byte1 = Chrb(((Char2 And 48) / 16) Or (Char1 * 4) And &HFF)
        Byte2 = lsGroupBinary & Chrb(((Char3 And 60) / 4) Or (Char2 * 16) And &HFF)
        Byte3 = Chrb((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
        
        If M4 = 2 Then
            lsGroupBinary = Byte1
        ElseIf M4 = 3 Then
            lsGroupBinary = Byte1 & Byte2
        End If
        
        lsResult = lsResult & lsGroupBinary
    End If
    
    Base64decode = lsResult
    
End Function

URL = inputbox("请输入您要转换的地址",,"http://www.yongfa365.com")
Msg=Msg & vbcrlf & "thunder://" & strAnsi2Unicode(Base64encode(strUnicode2Ansi("AA"&URL&"ZZ")))
Msg=Msg & vbcrlf & "flashget://" & strAnsi2Unicode(Base64encode(strUnicode2Ansi(URL)))
Msg=Msg & vbcrlf & "qqdl://" & strAnsi2Unicode(Base64encode(strUnicode2Ansi(URL)))
WScript.Echo Msg
