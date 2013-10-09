@echo off
title 柳永法专用垃圾清工具，正在清理中...
echo 正在清除系统垃圾文件,请稍等......
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
rem 上一行的操作会改变文件夹的属性，这个属性会影响asp+access程序,所以给注释掉了，换成下边的了,缺点是不能删除这个文件夹下的文件夹，但文件都删除了
del /f /s /q %windir%\temp\*.*
del /f /s /q %windir%\SoftwareDistribution\Download\*.*
del /f /s /q %windir%\ServicePackFiles\*.*
del /f /q %userprofile%\cookies\*.*
del /f /q %userprofile%\recent\*.*
del /f /s /q "%userprofile%\local settings\temporary internet files\*.*"
del /f /s /q "%userprofile%\local settings\temp\*.*"
del /f /s /q "%userprofile%\recent\*.*"
echo 清除系统垃圾文件完成!
echo. & pause