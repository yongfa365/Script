@echo off
title �Զ��ػ����� ����:�ལ
rem ����ĳ�������ֺ���
color 17
rem ����㲻ϲ��������Ĭ�ϵĺڵװ���ģʽ,������color������и���,����"17"�������װ���.
:start
echo.
echo.
echo ��ѡ��Ҫ���еĲ�����Ȼ�󰴻س���
echo.
echo 1. ��ʱ�ػ�
echo 2. ����ʱ�ػ�
echo 3. ɾ����ʱ�ػ�����
echo 4. �鿴��ʱ�ػ�����״̬
echo 5. ע��
echo 6. �˳�
echo. 
:set 
SET a=
SET /P a=ѡ��:
rem �趨����"a"Ϊ�û�������ַ�
IF NOT '%a%'=='' SET a=%a:~0,1%
ECHO.
IF /I '%a%'=='1' goto 1
IF /I '%a%'=='2' goto 2
IF /I '%a%'=='3' goto 3
IF /I '%a%'=='4' goto 4
IF /I '%a%'=='5' goto 5
IF /I '%a%'=='6' goto 6
rem ���������ַ�����1-6,��������������
echo %a% ѡ����Ч�����������룺
echo.
goto set
:1
echo ������ػ�ʱ��,(��12:00:00)
set shutdowntime=
set /p shutdowntime=
at %shutdowntime% tsshutdn 0 /delay:0 /powerdown >nul
IF not errorlevel 1 goto ok
rem ���������ȷ,��ִ��:ok��������
echo %shutdowntime% ���Ǳ�׼��ʱ���ʽ,����������
echo.
goto 1
:ok
echo.
echo �趨���! �����������...
pause >nul
cls
goto start
:2
echo ����Ҫ�������ػ�
echo (���趨��Ҫȡ��,����"ȷ��"��Ctrl+C������)
set timed=
set /p timed=����:
tsshutdn %timed% /delay:0 /powerdown >nul
IF not errorlevel 1 goto ok
echo %timed% ����Ч�Ĺػ�ʱ��,����������
echo.
goto 2
:3
at /del /y
echo ��ʱ�ػ�������ȡ��,�����������...
pause >nul
cls
goto start
:4
at
echo �����������...
pause >nul
cls
goto start
:5
logoff
:6
exit