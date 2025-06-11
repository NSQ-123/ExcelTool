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