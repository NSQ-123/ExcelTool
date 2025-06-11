@echo off  
chcp 65001 

REM 这些路径是相对于 .bat 文件所在位置的
:: 这些路径是相对于 .bat 文件所在位置的

set inputExcel=excel
set outputCsv=csvOutput
set outputCsharp=csharpOutput
set project=publish/ExcelTool.exe

REM 检查 inputExcel 是否存在
if not exist "%inputExcel%" (
    echo 错误：输入目录 "%inputExcel%" 不存在！
    pause
    exit /b 1
)

REM 检查并创建 outputCsv
if not exist "%outputCsv%" (
    echo 输出目录 "%outputCsv%" 不存在，正在创建...
    mkdir "%outputCsv%"
)

REM 检查并创建 outputCsharp
if not exist "%outputCsharp%" (
    echo 输出目录 "%outputCsharp%" 不存在，正在创建...
    mkdir "%outputCsharp%"
)



echo Running Program with:
for %%I in ("%inputExcel%") do echo Input Excel: %%~fI
for %%I in ("%outputCsv%") do echo Output CSV: %%~fI
for %%I in ("%outputCsharp%") do echo Output C#: %%~fI
for %%I in ("%project%") do echo project File: %%~fI


REM 运行源码
REM dotnet run -- %inputExcel% %outputCsv% %outputCsharp%  

REM 运行可执行文件
"%project%" "%inputExcel%" "%outputCsv%" "%outputCsharp%"

echo 运行完毕
pause

REM 常用修饰符
REM %%~fI：绝对路径（full path）
REM %%~dpI：驱动器和路径
REM %%~nxI：文件名和扩展名
REM %%~xI：扩展名
REM %%~nI：文件名（不含扩展名