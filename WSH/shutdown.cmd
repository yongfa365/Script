:start 
CLS 
@echo Off 
COLOR 1f 
rem 使用COLOR命令对控制台输出颜色进行更改 
MODE con: COLS=41 LINES=18 
rem MODE语句为设定窗体的宽和高 
set tm1=%time:~0,2%
set tm2=%time:~3,2%
set tm3=%time:~6,2%
ECHO %date% %tm1%点%tm2%分%tm3%秒 
ECHO ========================================= 
ECHO 请选择要进行的操作，然后按回车 
ECHO ─────────────── 
ECHO. 
ECHO 1. 定时关机 
ECHO 2. 倒计时关机 
ECHO 3. 删除定时关机任务 
ECHO 4. 查看任务状态 
ECHO 5. 刷新当前时间 
ECHO 6. 重新启动 
ECHO 7. 锁定计算机 
ECHO 8. 注销 
ECHO 9. 退出 
ECHO.
:cho 
SET Choice= 
SET /P Choice=选择: 
rem 设定变量"Choice"为用户输入的字符 
IF NOT "%Choice%"=="" SET Choice=%Choice:~0,1% 
rem 如果输入大于1位,取第1位,比如输入132,则返回值为1 
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
rem 为避免出现返回值为空或含空格而导致程序异常,需在变量外另加双引号 
rem 注意,IF语句需要双等于号 
rem 如果输入的字符不是以上数字,将返回重新输入 
ECHO 选择无效，请重新输入 
ECHO. 
GOTO cho
:SetHour 
CLS 
ECHO. 
SET ask= 
SET /p ask=是否设定为每天执行关机命令(y/n): 
IF NOT "%ask%"=="" SET ask=%ask:~0,1% 
IF /I "%ask%"=="y" GOTO yes 
IF /I "%ask%"=="n" GOTO no 
GOTO SetHour
:yes 
ECHO 请指定24小时制式时间,格式为 小时:分钟 
SET shutdowntime= 
SET /p shutdowntime=输入: 
at %shutdowntime% /every:M,T,W,Th,F,S,Su tsshutdn 0 /delay:0 /powerdown >nul 
rem 设定为每周的星期一至星期日,即为每天 
IF NOT errorlevel 1 GOTO ok 
rem 如果输入正确,就执行ok段的语句 
ECHO %shutdowntime% 不是标准的时间格式,请重新输入 
ECHO. 
GOTO yes
:no 
ECHO 请指定24小时制式时间,格式为 小时:分钟 
SET shutdowntime= 
SET /p shutdowntime=输入: 
at %shutdowntime% tsshutdn 0 /delay:0 /powerdown >nul 
IF NOT errorlevel 1 GOTO ok 
ECHO %shutdowntime% 不是标准的时间格式,请重新输入 
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
SET HM=%ah%时%am%分 
) ELSE ( 
SET HM=%bh%时%bm%分) 
rem 如果输入h:mm则HM=h时mm分,否则HM=hh时mm分 
IF /I "%ask%"=="y" ECHO 系统将于每天的%HM%关闭 
IF /I "%ask%"=="n" ECHO 系统将于%HM%关闭 
ECHO 设定完毕! 按任意键继续... 
PAUSE >nul 
GOTO start
:outtime 
CLS 
ECHO. 
ECHO 请输入倒计时秒数 
ECHO ──────── 
ECHO (设定后要取消,单击"确定"后按Ctrl+C键两次) 
SET timed= 
SET /p timed=输入: 
tsshutdn %timed% /delay:0 /powerdown >nul 
IF not errorlevel 1 GOTO ok 
ECHO %timed% 是无效的关机时间,请重新输入 
ECHO. 
GOTO outtime
:delAt 
cls 
echo. 
at /del /y 
echo 定时关机任务已取消,按任意键继续... 
pause >nul 
GOTO start
:view 
MODE con: COLS=85 LINES=18 
COLOR 70 
ECHO. 
at 
ECHO 按任意键继续... 
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
