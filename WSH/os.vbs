Set os=CreateObject("wscript.shell")
Set os0=CreateObject("shell.application")
Do 
input1=InputBox(" 请选择："+chr(13)+chr(13)+ _
        "1. 全部窗口最小化"+chr(13)+ _ 
        "2. 窗口状态复原"+chr(13)+ _ 
        "3. 窗口均匀平铺"+chr(13)+ _
        "4. 窗口纵向平铺"+chr(13)+ _
        "5. 窗口重叠展开" +Chr(13)+ _
        "6. 资源管理器"+chr(13)+ _ 
        "7. 设置系统时间" +Chr(13) + _
        "8. 刷新系统菜单" +Chr(13)+ _
        "9. 手动设置开始菜单"+Chr(13)+ _
        "10. 搜索文件"+Chr(13)+ _
        "11. 搜索计算机"+Chr(13)+ _
        "12. 运行"+Chr(13)+ _ 
        "13. 帮助"+Chr(13)+ _
        "14. 打开文件夹"+Chr(13)+ _
        "15. 挂起系统"+Chr(13)+ _
        "16. 关闭系统" +Chr(13)+ _
        "","vbs shell32 功能调用")
Select Case input1
Case 1
     os0.MinimizeAll
Case 2
     os0.UndoMinimizeALL
Case 3
     os0.TileHorizontally
Case 4
     os0.TileVertically
Case 5
     os0.CascadeWindows
Case 6
     p1=os.SpecialFolders("desktop")
     os0.Explore(p1)
Case 7
     os0.SetTime
Case 8
     os0.RefreshMenu
Case 9
     os0.TrayProperties
Case 10
     os0.FindFiles
Case 11
     os0.FindComputer
Case 12
     os0.FileRun
Case 13
     os0.Help
Case 14
     Set path1=os0.BrowseForFolder(0,"选择要打开的文件夹:",0)
     If path1 Is Nothing Then 
     Else
      os0.Open(path1.self.path)
     End If
Case 15
     os0.Suspend
Case 16 
     os0.ShutdownWindows
Case ""
     Exit Do
Case Else 
     os.Popup "选择错误",2,"错误",64+0
End Select
loop
