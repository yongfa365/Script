:start 
CLS 
@echo Off 
COLOR 1f 
rem ʹ��COLOR����Կ���̨�����ɫ���и��� 
MODE con: COLS=41 LINES=18 
rem MODE���Ϊ�趨����Ŀ�͸� 
set tm1=%time:~0,2%
set tm2=%time:~3,2%
set tm3=%time:~6,2%
ECHO %date% %tm1%��%tm2%��%tm3%�� 
ECHO ========================================= 
ECHO ��ѡ��Ҫ���еĲ�����Ȼ�󰴻س� 
ECHO ������������������������������ 
ECHO. 
ECHO 1. ��ʱ�ػ� 
ECHO 2. ����ʱ�ػ� 
ECHO 3. ɾ����ʱ�ػ����� 
ECHO 4. �鿴����״̬ 
ECHO 5. ˢ�µ�ǰʱ�� 
ECHO 6. �������� 
ECHO 7. ��������� 
ECHO 8. ע�� 
ECHO 9. �˳� 
ECHO.
:cho 
SET Choice= 
SET /P Choice=ѡ��: 
rem �趨����"Choice"Ϊ�û�������ַ� 
IF NOT "%Choice%"=="" SET Choice=%Choice:~0,1% 
rem ����������1λ,ȡ��1λ,��������132,�򷵻�ֵΪ1 
ECHO. 
IF /I "%Choice%"=="1" GOTO SetHour 
IF /I "%Choice%"=="2" GOTO outtime 
IF /I "%Choice%"=="3" GOTO delAt 
IF /I "%Choice%"=="4" GOTO view 
IF /I "%Choice%"=="5" GOTO start 
IF /I "%Choice%"=="6" GOTO restart 
IF /I "%Choice%"=="7" GOTO lock 
IF /I "%Choice%"=="8" GOTO logoff 
IF /I "%Choice%"=="9" GOTO end 
rem Ϊ������ַ���ֵΪ�ջ򺬿ո�����³����쳣,���ڱ��������˫���� 
rem ע��,IF�����Ҫ˫���ں� 
rem ���������ַ�������������,�������������� 
ECHO ѡ����Ч������������ 
ECHO. 
GOTO cho
:SetHour 
CLS 
ECHO. 
SET ask= 
SET /p ask=�Ƿ��趨Ϊÿ��ִ�йػ�����(y/n): 
IF NOT "%ask%"=="" SET ask=%ask:~0,1% 
IF /I "%ask%"=="y" GOTO yes 
IF /I "%ask%"=="n" GOTO no 
GOTO SetHour
:yes 
ECHO ��ָ��24Сʱ��ʽʱ��,��ʽΪ Сʱ:���� 
SET shutdowntime= 
SET /p shutdowntime=����: 
at %shutdowntime% /every:M,T,W,Th,F,S,Su tsshutdn 0 /delay:0 /powerdown >nul 
rem �趨Ϊÿ�ܵ�����һ��������,��Ϊÿ�� 
IF NOT errorlevel 1 GOTO ok 
rem ���������ȷ,��ִ��ok�ε���� 
ECHO %shutdowntime% ���Ǳ�׼��ʱ���ʽ,���������� 
ECHO. 
GOTO yes
:no 
ECHO ��ָ��24Сʱ��ʽʱ��,��ʽΪ Сʱ:���� 
SET shutdowntime= 
SET /p shutdowntime=����: 
at %shutdowntime% tsshutdn 0 /delay:0 /powerdown >nul 
IF NOT errorlevel 1 GOTO ok 
ECHO %shutdowntime% ���Ǳ�׼��ʱ���ʽ,���������� 
ECHO. 
GOTO no
:ok 
ECHO. 
SET h=%shutdowntime:~1,1% 
SET ah=%shutdowntime:~0,1% 
SET am=%shutdowntime:~2,2% 
SET bh=%shutdowntime:~0,2% 
SET bm=%shutdowntime:~3,2% 
IF "%h%"==":" ( 
SET HM=%ah%ʱ%am%�� 
) ELSE ( 
SET HM=%bh%ʱ%bm%��) 
rem �������h:mm��HM=hʱmm��,����HM=hhʱmm�� 
IF /I "%ask%"=="y" ECHO ϵͳ����ÿ���%HM%�ر� 
IF /I "%ask%"=="n" ECHO ϵͳ����%HM%�ر� 
ECHO �趨���! �����������... 
PAUSE >nul 
GOTO start
:outtime 
CLS 
ECHO. 
ECHO �����뵹��ʱ���� 
ECHO ���������������� 
ECHO (�趨��Ҫȡ��,����"ȷ��"��Ctrl+C������) 
SET timed= 
SET /p timed=����: 
tsshutdn %timed% /delay:0 /powerdown >nul 
IF not errorlevel 1 GOTO ok 
ECHO %timed% ����Ч�Ĺػ�ʱ��,���������� 
ECHO. 
GOTO outtime
:delAt 
cls 
echo. 
at /del /y 
echo ��ʱ�ػ�������ȡ��,�����������... 
pause >nul 
GOTO start
:view 
MODE con: COLS=85 LINES=18 
COLOR 70 
ECHO. 
at 
ECHO �����������... 
PAUSE >nul 
GOTO start
:restart 
shutdown -r -t 0
:lock 
rundll32.exe user32.dll,LockWorkStation 
goto start
:logoff 
logoff
:end 
exit 
