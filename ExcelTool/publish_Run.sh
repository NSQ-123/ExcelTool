#!/bin/bash

# 发布目录
publishDir="./publish"

# 发布项目
echo "Publishing project to $publishDir..."
dotnet publish -c Release -o "$publishDir"

# 检查发布是否成功
if [ $? -eq 0 ]; then
    echo "Publish succeeded. Executable is located in $publishDir"
else
    echo "Publish failed."
    exit 1
fi

# 运行发布后的程序
inputExcel="../../excel"
outputCsv="../../csvOutput"
outputCsharp="../../csharpOutput"

echo "Running published program with:"
echo "Input Excel: $inputExcel"
echo "Output CSV: $outputCsv"
echo "Output C#: $outputCsharp"

"$publishDir/ExcelTool" "$inputExcel" "$outputCsv" "$outputCsharp"