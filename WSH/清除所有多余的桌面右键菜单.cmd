@ ECHO OFF
@ ECHO.
@ ECHO.
@ ECHO.                              ˵  ��
@ ECHO -----------------------------------------------------------------------
@ ECHO �ܶ��Կ���װ������֮�������Ҽ�����һ������˵�����Щ���ܲ���ʵ�ã�
@ ECHO ���������Ҽ��ĵ����ٶȣ���Intel�ļ����Կ�Ϊ�����ٴ��ķ�Ӧ�ٶ����ص�Ӱ��
@ ECHO ��ʹ���ߵ����顣����������������GhostXP���Թ�˾�ر�桷���߱ࡣ
@ ECHO -----------------------------------------------------------------------
PAUSE

regsvr32 /u /s igfxpph.dll
reg delete HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers /f
reg add HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers\new /ve /d {D969A300-E7FF-11d0-A93B-00A0C90F2719}
reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v HotKeysCmds /f
reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v IgfxTray /f
