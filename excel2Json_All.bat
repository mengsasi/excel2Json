@echo off & setlocal EnableDelayedExpansion

::导出目录
set exportPath="configs\\"

set cur=%cd%"\configs"
echo %cur%
::if exist %cur% rd /s/q %cur%
::删除旧文件

python excel2Json_All.py %exportPath%

pause