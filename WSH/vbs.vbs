Function Decrypt(password)
    magic = Split("121,65,51,54,122,65,52,56,100,69,104,102,114,118,103,104,71,82,103,53,55,104,53,85,108,68,118,51", ",")
    chrlast = CInt("&H" & Mid(password, 1, 2))
    magicnum = 0
    
    For X = 3 To Len(password) Step 2
        chrtmp = CInt("&H" & Mid(password, X, 2))
        chrresulta = (chrtmp Xor magic(magicnum))
        chrresultb = chrresulta - CInt(chrlast)
        
        If chrresultb > 255 Or chrresultb < 0 Then
            chrresultb = chrresultb - &HFFFFFF01
        End If
        chrlast = chrtmp
        pwdtmp = pwdtmp & Chr(chrresultb)
        magicnum = magicnum + 1
        
        If magicnum > 27 Then
            magicnum = 0
        End If
    Next
    
    Decrypt = pwdtmp
End Function

MsgBox Decrypt("41F072E8799083F973B8BF99987D81886A")
