@echo off
chcp 65001 

REM 发布目录
set publishDir=publish

for %%I in ("%publishDir%") do echo publish Dir: %%~fI
REM 发布项目
echo Publishing project to %publishDir%...
dotnet publish -c Release -o %publishDir%

if %ERRORLEVEL% NEQ 0 (
    echo Publish failed.
    exit /b 1
)

pause