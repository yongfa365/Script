Const ADMINISTRATIVE_TOOLS = &H2f&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS) 
Set objTools = objFolder.Items
For i = 0 to objTools.Count - 1
    tools=tools&objTools.Item(i)&vbcrlf
Next
wscript.echo "有如下管理工具："&vbcrlf&tools
set objShell=nothing
set objFolder=nothing
set objTools=nothing