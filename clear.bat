@echo off
title ������ר�������幤�ߣ�����������...
echo �������ϵͳ�����ļ�,���Ե�......
del /f /s /q %systemdrive%\*.tmp
del /f /s /q %systemdrive%\*._mp
del /f /s /q %systemdrive%\*.log
del /f /s /q %systemdrive%\*.gid
del /f /s /q %systemdrive%\*.chk
del /f /s /q %systemdrive%\*.old
del /f /s /q %systemdrive%\recycled\*.*
del /f /s /q %systemdrive%\*.sqm
del /f /s /q %windir%\*.bak
del /f /s /q %windir%\prefetch\*.*
rem rd /s /q %windir%\temp & md %windir%\temp
rem ��һ�еĲ�����ı��ļ��е����ԣ�������Ի�Ӱ��asp+access����,���Ը�ע�͵��ˣ������±ߵ���,ȱ���ǲ���ɾ������ļ����µ��ļ��У����ļ���ɾ����
del /f /s /q %windir%\temp\*.*
del /f /s /q %windir%\SoftwareDistribution\Download\*.*
del /f /s /q %windir%\ServicePackFiles\*.*
del /f /q %userprofile%\cookies\*.*
del /f /q %userprofile%\recent\*.*
del /f /s /q "%userprofile%\local settings\temporary internet files\*.*"
del /f /s /q "%userprofile%\local settings\temp\*.*"
del /f /s /q "%userprofile%\recent\*.*"
echo ���ϵͳ�����ļ����!
echo. & pause