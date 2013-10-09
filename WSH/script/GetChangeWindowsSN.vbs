'/*=========================================================================
' * Intro       查看或修改Windows系统的序列号，支持命令行或直接运行输入。
' * FileName    GetChangeWindowsSN.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/Item/GetChangeWindowsSN.vbs.html
' * MadeTime    2007-10-13 21:40:09
' * LastModify  2007-10-13 21:40:09
' *==========================================================================*/

On Error Resume Next
SN_XP_1 = "MRX3F-47B9T-2487J-KWKMF-RPWBY" 'good
SN_XP_2 = "QC986-27D34-6M3TY-JJXP9-TBGMD"
SN_XP_3 = "K2CXT-C6TPX-WCXDP-RMHWT-V4TDT"
SN_XP_4 = "22DVC-GWQW7-7G228-D72Y7-QK8Q3"
SN_XP_5 = "DG8FV-B9TKY-FRT9J-6CRCC-XPQ4G"
SN_XP_6 = "T44H2-BM3G7-J4CQR-MPDRM-BWFWM"
SN_XP_7 = "XW6Q2-MP4HK-GXFK3-KPGG4-GM36T"

SN_2000_1 = "PQHKR-G4JFW-VTY3P-G4WQ2-88CTW"
SN_2000_Server_1 = "H6TWQ-TQQM8-HXJYG-D69F7-R84VM"
SN_2000_Advanced_Server_1 = "H6TWQ-TQQM8-HXJYG-D69F7-R84VM"

SN_2003_1 = "JCGMJ-TC669-KCBG7-HB8X2-FXG7M" 'good
SN_2003_2 = "DF74D-TWR86-D3F4V-M8D8J-WTT7M" 'good
SN_2003_2 = "KQF2H-284RW-GHXM6-Y3W2B-QWGBB"


Dim VOL_PROD_KEY
If WScript.arguments.Count<1 Then
    VOL_PROD_KEY = InputBox("您当前的Windows系统序列号为：" & GetWindowsSN & String(5, vbCrLf) & "请输入新的Windows序列号：", "Windows序列号更换器", SN_2003_1)
    If VOL_PROD_KEY = "" Or Len(VOL_PROD_KEY)<>29 Then
        WScript.echo "您选择了取消 或 Windows序列号为空 或 Windows序列号位数有误  ——》退出"
        WScript.Quit
    End If
Else
    VOL_PROD_KEY = WScript.arguments.Item(0)
End If

VOL_PROD_KEY = Replace(VOL_PROD_KEY, "-", "") 'remove hyphens if any

For Each Obj in GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf ("win32_WindowsProductActivation")
    result = Obj.SetProductKey (VOL_PROD_KEY)
    If Err = 0 Then
        WScript.echo "Windows序列号替换成功。"
    Else
        WScript.echo "Windows序列号替换失败！您输入的序列号有误。"
        Err.Clear
    End If
Next


'取得当前Windows序列号函数

Function GetWindowsSN()
    Const HKEY_LOCAL_MACHINE = &H80000002
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    strValueName = "DigitalProductId"
    strComputer = "."
    Dim iValues()
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    oReg.GetBinaryValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, iValues
    Dim arrDPID
    arrDPID = Array()
    For i = 52 To 66
        ReDim Preserve arrDPID( UBound(arrDPID) + 1 )
        arrDPID( UBound(arrDPID) ) = iValues(i)
    Next
    ' <--------------- Create an array to hold the valid characters for a microsoft Product Key -------------------------->
    Dim arrChars
    arrChars = Array("B", "C", "D", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "6", "7", "8", "9")
    
    ' <--------------- The clever bit !!! (Decrypt the base24 encoded binary data)-------------------------->
    For i = 24 To 0 Step -1
        k = 0
        For j = 14 To 0 Step -1
            k = k * 256 Xor arrDPID(j)
            arrDPID(j) = Int(k / 24)
            k = k Mod 24
        Next
        strProductKey = arrChars(k) & strProductKey
        ' <------- add the "-" between the groups of 5 Char -------->
        If i Mod 5 = 0 And i <> 0 Then strProductKey = "-" & strProductKey
    Next
    GetWindowsSN = strProductKey
End Function
