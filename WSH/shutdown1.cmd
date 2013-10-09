@echo off
title 自动关机程序 作者:青剑
rem 这里改成你的名字好了
color 17
rem 如果你不喜欢命令行默认的黑底白字模式,可以用color命令进行更改,上面"17"代表蓝底白字.
:start
echo.
echo.
echo 请选择要进行的操作，然后按回车：
echo.
echo 1. 定时关机
echo 2. 倒计时关机
echo 3. 删除定时关机任务
echo 4. 查看定时关机任务状态
echo 5. 注销
echo 6. 退出
echo. 
:set 
SET a=
SET /P a=选择:
rem 设定变量"a"为用户输入的字符
IF NOT '%a%'=='' SET a=%a:~0,1%
ECHO.
IF /I '%a%'=='1' goto 1
IF /I '%a%'=='2' goto 2
IF /I '%a%'=='3' goto 3
IF /I '%a%'=='4' goto 4
IF /I '%a%'=='5' goto 5
IF /I '%a%'=='6' goto 6
rem 如果输入的字符不是1-6,将返回重新输入
echo %a% 选择无效，请重新输入：
echo.
goto set
:1
echo 请输入关机时间,(如12:00:00)
set shutdowntime=
set /p shutdowntime=
at %shutdowntime% tsshutdn 0 /delay:0 /powerdown >nul
IF not errorlevel 1 goto ok
rem 如果输入正确,就执行:ok后面的语句
echo %shutdowntime% 不是标准的时间格式,请重新输入
echo.
goto 1
:ok
echo.
echo 设定完毕! 按任意键继续...
pause >nul
cls
goto start
:2
echo 您想要多少秒后关机
echo (若设定后要取消,单击"确定"后按Ctrl+C键两次)
set timed=
set /p timed=输入:
tsshutdn %timed% /delay:0 /powerdown >nul
IF not errorlevel 1 goto ok
echo %timed% 是无效的关机时间,请重新输入
echo.
goto 2
:3
at /del /y
echo 定时关机任务已取消,按任意键继续...
pause >nul
cls
goto start
:4
at
echo 按任意键继续...
pause >nul
cls
goto start
:5
logoff
:6
exit