Set os=CreateObject("wscript.shell")
Set os0=CreateObject("shell.application")
Do 
input1=InputBox(" ��ѡ��"+chr(13)+chr(13)+ _
        "1. ȫ��������С��"+chr(13)+ _ 
        "2. ����״̬��ԭ"+chr(13)+ _ 
        "3. ���ھ���ƽ��"+chr(13)+ _
        "4. ��������ƽ��"+chr(13)+ _
        "5. �����ص�չ��" +Chr(13)+ _
        "6. ��Դ������"+chr(13)+ _ 
        "7. ����ϵͳʱ��" +Chr(13) + _
        "8. ˢ��ϵͳ�˵�" +Chr(13)+ _
        "9. �ֶ����ÿ�ʼ�˵�"+Chr(13)+ _
        "10. �����ļ�"+Chr(13)+ _
        "11. ���������"+Chr(13)+ _
        "12. ����"+Chr(13)+ _ 
        "13. ����"+Chr(13)+ _
        "14. ���ļ���"+Chr(13)+ _
        "15. ����ϵͳ"+Chr(13)+ _
        "16. �ر�ϵͳ" +Chr(13)+ _
        "","vbs shell32 ���ܵ���")
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
     Set path1=os0.BrowseForFolder(0,"ѡ��Ҫ�򿪵��ļ���:",0)
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
     os.Popup "ѡ�����",2,"����",64+0
End Select
loop
