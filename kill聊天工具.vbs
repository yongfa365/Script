Set wmi = GetObject("winmgmts:\\.")
Set pro_s = wmi.instancesof("win32_process")

For Each p In pro_s
    If LCase(p.Name) = "qq.exe" Then p.terminate()
    If LCase(p.Name) = "tm.exe" Then p.terminate()
    If LCase(p.Name) = "fetion.exe" Then p.terminate()
Next
