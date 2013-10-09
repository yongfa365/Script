Dim fso,file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.GetFile(Wscript.ScriptFullName)
file.Copy "C:\"