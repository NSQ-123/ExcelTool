@echo off
//这些路径是相对于 .bat 文件所在位置的
set inputExcel=..\..\excel
set outputCsv=..\..\csvOutput
set outputCsharp=..\..\csharpOutput

echo Running Program with:
echo Input Excel: %inputExcel%
echo Output CSV: %outputCsv%
echo Output C#: %outputCsharp%

dotnet run -- %inputExcel% %outputCsv% %outputCsharp%

pause