@echo off & setlocal EnableDelayedExpansion

::excel表目录
set excelPath=""
::导出目录
set exportPath=""

set cur=%cd%"\configs"
echo %cur%
if exist %cur% rd /s/q %cur%
::删除旧文件

python excel2Json_All.py

pause