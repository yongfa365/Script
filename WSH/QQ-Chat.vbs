'/*=========================================================================
' * Intro       ͨ����Ѷ��ʱ�Ự����ʵ�֣�QQǿ�����칤��
' * FileName    QQ-Chat.vbs
' * Author      yongfa365
' * Version     v1.0
' * WEB         http://www.yongfa365.com
' * Email       yongfa365[at]qq.com
' * FirstWrite  http://www.yongfa365.com/blog/item/QQ-QiangXingLiaoTianGongJu.htm
' * LastModify  2007-08-31 15:11:50
' *==========================================================================*/

On Error Resume Next
num = InputBox("������Է�QQ����:", "ǿ�����칤��(����)") '����Է�QQ����
If num = "" Then WScript.Quit '�������ı���Ϊ�����������
url = "tencent://message/?uin=" + num + "&Site=1&Menu=yes" '��Ѷ��ʱ�Ự����
Set obj = WScript.CreateObject("InternetExplorer.Application") '����IE���������
obj.Visible = False '����IE�����Ϊ���ɼ�
obj.Navigate(url) '��IE����� ��Ч��Ϊ��QQ��ʱ�Ự����

