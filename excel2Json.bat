@echo off & setlocal EnableDelayedExpansion

::excel表目录 上级目录 ..\\..\\test\\
set excelPath="test\\"
::导出目录
set exportPath="configs\\"

set cur=%cd%"\configs"
::echo %cur%
::if exist %cur% rd /s/q %cur%
::删除旧文件

for /f "delims=" %%i in ('dir /a-d /b *.xlsx') do (  
    ::echo %%i
	python excel2Json.py %excelPath%%%i %exportPath%
)

pause