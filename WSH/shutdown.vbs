'==========================================================================
'
' ע��/����/�رձ���Windows NT/2000 �����������˼·���£�
'
' Win32ShutDown(flag)��flag�Ĳ���:
' 0 ע��
' 0 + 4 ǿ��ע��
' 1 �ػ�
' 1 + 4 ǿ�ƹػ�
' 2 ����
' 2 + 4 ǿ������
' 8 �رյ�Դ
' 8 + 4 ǿ�ƹرյ�Դ
'
'==========================================================================

ShutDown

Sub ShutDown()
    Dim Connection, WQL, SystemClass, System
    Set Connection = GetObject("winmgmts:root\cimv2")
    WQL = "Select Name From Win32_OperatingSystem"
    Set SystemClass = Connection.ExecQuery(WQL)
    For Each System In SystemClass
        System.Win32ShutDown (2 + 4)
    Next
End Sub
