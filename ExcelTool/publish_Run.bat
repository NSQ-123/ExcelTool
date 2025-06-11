@echo off

REM 发布目录
set publishDir=publish

REM 发布项目
echo Publishing project to %publishDir%...
dotnet publish -c Release -o %publishDir%

if %ERRORLEVEL% NEQ 0 (
    echo Publish failed.
    exit /b 1
)

REM 运行发布后的程序
set inputExcel=..\..\excel
set outputCsv=..\..\csvOutput
set outputCsharp=..\..\csharpOutput

echo Running published program with:
echo Input Excel: %inputExcel%
echo Output CSV: %outputCsv%
echo Output C#: %outputCsharp%

%publishDir%\ExcelTool.exe %inputExcel% %outputCsv% %outputCsharp%