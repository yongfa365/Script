@echo off
echo �������ϵͳ�����ļ�,���Ե�......
del /f /s /q %systemdrive%\*.tmp
del /f /s /q %systemdrive%\*._mp
del /f /s /q %systemdrive%\*.gid
del /f /s /q %systemdrive%\*.chk
del /f /s /q %systemdrive%\*.old
del /f /s /q %windir%\*.bak
del /f /s /q %windir%\temp\*.*
del /f /a /q %systemdrive%\*.sqm
del /f /s /q %windir%\SoftwareDistribution\Download\*.*
del /f /s /q "%userprofile%\cookies\*.*"
del /f /s /q "%userprofile%\recent\*.*"
del /f /s /q "%userprofile%\local settings\temporary internet files\*.*"
del /f /s /q "%userprofile%\local settings\temp\*.*"
echo ���ϵͳ�����ļ����!
echo. & pause