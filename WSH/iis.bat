@ECHO OFF
rem setlocal

REM =======================================
REM ��������C��Ȩ������ [IIS]  V1.3
REM v1.3, C.Rufus Security Team PSD
REM =======================================
REM
REM CHANGELOG --
REM by amxku&�����ֻ�, C.Rufus S.T
REM 2006-12-10

REM add some tips ;)
REM by amxku, C.Rufus S.T
REM 2007-07-10

title ��������C��Ȩ������ [IIS]  V1.3 - ���ǰ�ȫС��
echo.
echo "++++++++++++++++++++++++++++++++++++++++++++++++++++"
echo "+          ��������C��Ȩ������ [IIS]  V1.3          "
echo "+                                                   "
echo "+               Www.wolfexp.net                     "
echo "+                ���ǰ�ȫС��                       "
echo "+                                                   "
echo "+              amxku   �����ֻ�                     "
echo "++++++++++++++++++++++++++++++++++++++++++++++++++++"
:menu
echo.
echo [1]     ɾ��C�̵�everyone��Ȩ��
echo [2]     ɾ��C�̵����е�users�ķ���Ȩ��
echo [3]     ���iis_wpg�ķ���Ȩ��
echo [4]     ���iis_wpg�ķ���Ȩ��[.netר��]
echo [5]     ���iis_wpg�ķ���Ȩ��[װ��MACFEE�����ר��]
echo [6]     ���users�ķ���Ȩ��
echo [7]     ɾ��C��Windows�µ����е�Σ���ļ���
echo [8]     ɾ��ϵͳΣ���ļ��ķ���Ȩ�ޣ�ֻ���������Ա
echo [9]     ע�������趨
echo [0]     �˳�
echo.
@echo ��ѡ��?
@echo ���������ѡ��س�
@echo off
set /p menu=

if %menu% == 0 goto exit
if %menu% == 1 goto 1
if %menu% == 2 goto 2
if %menu% == 3 goto 3
if %menu% == 4 goto 4
if %menu% == 5 goto 5
if %menu% == 6 goto 6
if %menu% == 7 goto 7
if %menu% == 8 goto 8
if %menu% == 9 goto 9

:1
echo ɾ��C�̵�everyone��Ȩ�� 
cacls "%SystemDrive%" /r "CREATOR OWNER" /e
cacls "%SystemDrive%" /r "everyone" /e 
cacls "%SystemRoot%" /r "everyone" /e 
cacls "%SystemDrive%/Documents and Settings" /r "everyone" /e 
cacls "%SystemDrive%/Documents and Settings/All Users" /r "everyone" /e 
cacls "%SystemDrive%/Documents and Settings/All Users/Documents"  /r "everyone" /e 
echo.
echo ɾ��C�̵�everyone��Ȩ�� ������������ok!
echo.
goto menu

:2
echo ɾ��C�̵����е�users�ķ���Ȩ�� 
cacls "%SystemDrive%" /r "users" /e 
cacls "%SystemDrive%/Program Files" /r "users" /e 
cacls "%SystemDrive%/Documents and Settings" /r "users" /e 
cacls "%SystemRoot%" /r "users" /e 
cacls "%SystemRoot%/addins" /r "users" /e 
cacls "%SystemRoot%/AppPatch" /r "users" /e 
cacls "%SystemRoot%/Connection Wizard" /r "users" /e 
cacls "%SystemRoot%/Debug" /r "users" /e 
cacls "%SystemRoot%/Driver Cache" /r "users" /e 
cacls "%SystemRoot%/Help" /r "users" /e 
cacls "%SystemRoot%/IIS Temporary Compressed Files" /r "users" /e 
cacls "%SystemRoot%/java" /r "users" /e 
cacls "%SystemRoot%/msagent" /r "users" /e 
cacls "%SystemRoot%/mui" /r "users" /e 
cacls "%SystemRoot%/repair" /r "users" /e 
cacls "%SystemRoot%/Resources" /r "users" /e 
cacls "%SystemRoot%/security" /r "users" /e 
cacls "%SystemRoot%/system" /r "users" /e 
cacls "%SystemRoot%/TAPI" /r "users" /e 
cacls "%SystemRoot%/Temp" /r "users" /e 
cacls "%SystemRoot%/twain_32" /r "users" /e 
cacls "%SystemRoot%/Web" /r "users" /e 
cacls "%SystemRoot%/WinSxS" /r "users" /e 
cacls "%SystemRoot%/system32/3com_dmi" /r "users" /e 
cacls "%SystemRoot%/system32/administration" /r "users" /e 
cacls "%SystemRoot%/system32/Cache" /r "users" /e 
cacls "%SystemRoot%/system32/CatRoot2" /r "users" /e 
cacls "%SystemRoot%/system32/Com" /r "users" /e 
cacls "%SystemRoot%/system32/config" /r "users" /e 
cacls "%SystemRoot%/system32/dhcp" /r "users" /e 
cacls "%SystemRoot%/system32/drivers" /r "users" /e 
cacls "%SystemRoot%/system32/export" /r "users" /e 
cacls "%SystemRoot%/system32/icsxml" /r "users" /e 
cacls "%SystemRoot%/system32/lls" /r "users" /e 
cacls "%SystemRoot%/system32/LogFiles" /r "users" /e 
cacls "%SystemRoot%/system32/MicrosoftPassport" /r "users" /e 
cacls "%SystemRoot%/system32/mui" /r "users" /e 
cacls "%SystemRoot%/system32/oobe" /r "users" /e 
cacls "%SystemRoot%/system32/ShellExt" /r "users" /e 
cacls "%SystemRoot%/system32/wbem" /r "users" /e 
echo.
echo ɾ��C�̵����е�users�ķ���Ȩ��  ������������ok!
echo.
goto menu


:7
echo ɾ��C��Windows�µ����е�Σ���ļ��� 
attrib %SystemRoot%/Web/printers -s -r -h
del %SystemRoot%\Web\printers\*.* /s /q /f
rd %SystemRoot%\Web\printers /s /q


attrib %SystemRoot%\system32\inetsrv\iisadmpwd -s -r -h
del %SystemRoot%\system32\inetsrv\iisadmpwd\*.* /s /q /f
rd %SystemRoot%\system32\inetsrv\iisadmpwd /s /q
echo.
echo ɾ��C��Windows�µ����е�Σ���ļ���   ������������ok!
echo.
goto menu


:8
echo ��ϵͳΣ���ļ�����Ȩ���趨
cacls "C:\boot.ini" /T /C /E /G Administrators:F
cacls "C:\boot.ini" /D Guests:F /E

cacls "C:\AUTOEXEC.BAT" /T /C /E /G Administrators:F
cacls "C:\AUTOEXEC.BAT" /D Guests:F /E

cacls "%SystemRoot%/system32/net.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/net.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/net1.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/net1.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/cmd.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/cmd.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/ftp.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/ftp.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/netstat.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/netstat.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/regedit.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/regedit.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/at.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/at.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/attrib.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/attrib.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/format.com" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/format.com" /D Guests:F /E

cacls "%SystemRoot%/system32/logoff.exe" /T /C /E /G Administrators:F

cacls "%SystemRoot%/system32/shutdown.exe" /G Administrators:F
cacls "%SystemRoot%/system32/shutdown.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/telnet.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/telnet.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/wscript.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/wscript.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/doskey.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/doskey.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/help.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/help.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/ipconfig.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/ipconfig.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/nbtstat.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/nbtstat.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/print.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/print.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/xcopy.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/xcopy.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/edit.com" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/edit.com" /D Guests:F /E

cacls "%SystemRoot%/system32/regedt32.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/regedt32.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/reg.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/reg.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/register.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/register.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/replace.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/replace.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/nwscript.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/nwscript.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/share.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/share.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/ping.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/ping.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/ipsec6.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/ipsec6.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/netsh.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/netsh.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/debug.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/debug.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/route.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/route.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/tracert.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/tracert.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/powercfg.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/powercfg.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/nslookup.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/nslookup.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/arp.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/arp.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/rsh.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/rsh.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/netdde.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/netdde.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/mshta.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/mshta.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/mountvol.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/mountvol.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/tftp.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/tftp.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/setx.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/setx.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/find.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/find.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/finger.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/finger.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/where.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/where.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/regsvr32.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/regsvr32.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/cacls.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/cacls.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/sc.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/sc.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/shadow.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/shadow.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/runas.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/runas.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/wshom.ocx" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/wshom.ocx" /D Guests:F /E

cacls "%SystemRoot%/system32/wshext.dll" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/wshext.dll" /D Guests:F /E

cacls "%SystemRoot%/system32/shell32.dll" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/shell32.dll" /D Guests:F /E

cacls "%SystemRoot%/system32/zipfldr.dll" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/zipfldr.dll" /D Guests:F /E

cacls "%SystemRoot%/PCHealth/HelpCtr/Binaries/msconfig.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/PCHealth/HelpCtr/Binaries/msconfig.exe" /D Guests:F /E

cacls "%SystemRoot%/notepad.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/notepad.exe" /D Guests:F /E

cacls "%SystemRoot%/regedit.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/regedit.exe" /D Guests:F /E

cacls "%SystemRoot%/winhelp.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/winhelp.exe" /D Guests:F /E

cacls "%SystemRoot%/winhlp32.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/winhlp32.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/notepad.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/notepad.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/edlin.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/edlin.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/posix.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/posix.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/atsvc.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/atsvc.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/qbasic.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/qbasic.exe" /T /C /E /G Administrators:F

cacls "%SystemRoot%/system32/runonce.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/runonce.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/syskey.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/syskey.exe" /D Guests:F /E

cacls "%SystemRoot%/system32/cscript.exe" /T /C /E /G Administrators:F
cacls "%SystemRoot%/system32/cscript.exe" /D Guests:F /E
echo.
echo ��ϵͳΣ���ļ�����Ȩ���趨   ������������ok!
echo.
goto menu

:9
echo ע�������趨
reg delete HKEY_CLASSES_ROOT\WScript.Shell /f
reg delete HKEY_CLASSES_ROOT\WScript.Shell.1 /f
reg delete HKEY_CLASSES_ROOT\Shell.application /f
reg delete HKEY_CLASSES_ROOT\Shell.application.1 /f
reg delete HKEY_CLASSES_ROOT\WSCRIPT.NETWORK /f
reg delete HKEY_CLASSES_ROOT\WSCRIPT.NETWORK.1 /f
regsvr32 /s /u wshom.ocx
regsvr32 /s /u wshext.dll
regsvr32 /s /u shell32.dll
regsvr32 /s /u zipfldr.dll
echo.
echo ע�������趨   ������������ok!
echo.
goto menu


:3
echo ���iis_wpg�ķ���Ȩ�� 
cacls "%SystemRoot%" /g iis_wpg:r /e 
cacls "%SystemDrive%/Program Files/Common Files" /g iis_wpg:r /e 

cacls "%SystemRoot%/Downloaded Program Files" /g iis_wpg:c /e 
cacls "%SystemRoot%/Help" /g iis_wpg:c /e 
cacls "%SystemRoot%/IIS Temporary Compressed Files" /g iis_wpg:c /e 
cacls "%SystemRoot%/Offline Web Pages" /g iis_wpg:c /e 
cacls "%SystemRoot%/System32" /g iis_wpg:c /e 
cacls "%SystemRoot%/Tasks" /g iis_wpg:c /e 
cacls "%SystemRoot%/Temp" /g iis_wpg:c /e 
cacls "%SystemRoot%/Web" /g iis_wpg:c /e 
echo.
echo ���iis_wpg�ķ���Ȩ��   ������������ok!
echo.
goto menu


:4
echo ���iis_wpg�ķ���Ȩ��[.netר��] 
cacls "%SystemRoot%/Assembly" /g iis_wpg:c /e 
cacls "%SystemRoot%/Microsoft.NET" /g iis_wpg:c /e 
echo.
echo ���iis_wpg�ķ���Ȩ��[.netר��]   ������������ok!
echo.
goto menu

:5
echo ���iis_wpg�ķ���Ȩ��[װ��MACFEE�����ר��] 
cacls "%SystemDrive%/Program Files/Network Associates" /g iis_wpg:r /e 
echo.
echo ���iis_wpg�ķ���Ȩ��[װ��MACFEE�����ר��]   ������������ok!
echo.
goto menu

:6
echo ���users�ķ���Ȩ�� 
cacls "%SystemRoot%/temp" /g users:c /e 
echo.
echo ���users�ķ���Ȩ��   ������������ok!
echo.
goto menu

:exit

exit