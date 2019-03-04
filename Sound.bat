@echo off & setlocal EnableDelayedExpansion

::excel表目录
set excelPath=""
::导出目录
set exportPath=""

set file=%~n0%.xlsx
::删除旧文件
set existFile=%exportPath%\%file%
if exist %existFile% rd /s/q %existFile%

if exist %file% python excel2Json.py %file%

pause