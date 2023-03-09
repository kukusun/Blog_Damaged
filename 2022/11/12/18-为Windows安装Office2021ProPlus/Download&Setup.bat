%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

@echo off
set kmsserver=192.168.100.220
title Office下载安装
:start
cls
echo 1.下载 Office
echo 2.安装 Office
echo 3.激活 Office2021ProPlus 版
echo 4.查询 Office 激活
echo 5.安装 VOL 许可证
set /p input=请选择：
if "%input%"=="1" goto 1
if "%input%"=="2" goto 2
if "%input%"=="3" goto 3
if "%input%"=="4" goto 4
if "%input%"=="5" goto 5
if not defined "%input%" goto start

:1
cd /d %~dp0
echo ----------------
echo 正在下载Office
setup.exe /download "configuration-2021.xml"
echo 下载完成
echo ----------------
pause
goto start

:2
cd /d %~dp0
echo ----------------
echo 正在安装Office
setup.exe /configure "configuration-2021.xml"
echo 安装完成
echo ----------------
pause
goto start

:3
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /sethst:%kmsserver%
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /act
pause
goto start

:4
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /dstatus
pause
goto start

:5
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /inslic:"%PROGRAMFILES%\Microsoft Office\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms"
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /inslic:"%PROGRAMFILES%\Microsoft Office\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul.xrm-ms"
cscript "%PROGRAMFILES%\Microsoft Office\Office16\OSPP.vbs" /inslic:"%PROGRAMFILES%\Microsoft Office\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul-oob.xrm-ms"
pause
goto start