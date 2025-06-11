#!/bin/bash

# 这些路径是相对于 .sh 文件所在位置的
inputExcel="../../excel"
outputCsv="../../csvOutput"
outputCsharp="../../csharpOutput"

echo "Running Program with:"
echo "Input Excel: $inputExcel"
echo "Output CSV: $outputCsv"
echo "Output C#: $outputCsharp"

# 运行程序并传递参数
dotnet run -- "$inputExcel" "$outputCsv" "$outputCsharp"