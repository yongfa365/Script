Set wmi=GetObject("winmgmts:\\.")
Set pro_s=wmi.instancesof("win32_process")

For Each p In pro_s
if p.name="wscript.exe" then p.terminate()
Next
