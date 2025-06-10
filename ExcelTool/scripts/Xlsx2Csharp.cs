using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口

public class Xlsx2Csharp
{
    /// <summary>
    /// 将 Excel 文件的第二个工作表转换为 C# 类定义
    /// </summary>
    /// <param name="excelFilePath">Excel 文件路径</param>
    /// <param name="outputFilePath">生成的 C# 文件路径</param>
    public static void ConvertToCsharp(string excelFilePath, string outputFilePath)
    {
        // 打开 Excel 文件
        using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(1); // 获取第二个工作表

            // 获取文件名称作为类名
            var fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            string className = "T_" + fileName;

            // 创建 StringBuilder 用于生成 C# 类定义
            StringBuilder classBuilder = new StringBuilder();
            classBuilder.AppendLine("using System;");
            classBuilder.AppendLine("using System.Collections.Generic;");
            classBuilder.AppendLine();
            classBuilder.AppendLine($"public partial class {className}");
            classBuilder.AppendLine("{");

            // 添加静态字典字段
            classBuilder.AppendLine($"    private static Dictionary<int, {className}> _dataDic = new Dictionary<int, {className}>();");
            classBuilder.AppendLine($"    private static List<{className}> _dataList;");
            classBuilder.AppendLine();

            // 遍历字段定义
            IRow fieldNameRow = sheet.GetRow(0); // 第一行：字段名称
            IRow fieldTypeRow = sheet.GetRow(1); // 第二行：字段类型
            IRow usageRow = sheet.GetRow(2);     // 第三行：使用方
            IRow descriptionRow = sheet.GetRow(3); // 第四行：字段描述

            if (fieldNameRow == null || fieldTypeRow == null || usageRow == null || descriptionRow == null)
            {
                throw new Exception("工作表格式不正确，缺少必要的字段定义行。");
            }

            StringBuilder fieldLoadBuilder = new StringBuilder();
            for (int i = 0; i < fieldNameRow.LastCellNum; i++)
            {
                string fieldName = fieldNameRow.GetCell(i)?.StringCellValue ?? "";
                string fieldType = fieldTypeRow.GetCell(i)?.StringCellValue ?? "string";
                string usage = usageRow.GetCell(i)?.StringCellValue ?? "";
                string description = descriptionRow.GetCell(i)?.StringCellValue ?? "";

                // 仅生成客户端使用的字段（含有 "c"）
                if (string.IsNullOrEmpty(usage))
                {
                    continue; // 如果使用方为空，则跳过该字段
                }
                usage = usage.ToLowerInvariant(); // 转为小写以便比较
                if (usage.Contains("c"))
                {
                    // 添加字段描述作为注释
                    if (!string.IsNullOrWhiteSpace(description))
                    {
                        classBuilder.AppendLine($"    /// <summary>");
                        classBuilder.AppendLine($"    /// {description}");
                        classBuilder.AppendLine($"    /// </summary>");
                    }

                    // 添加字段定义
                    fieldType = ConvertUtils.GetType(fieldType,fieldName); // 规范化字段类型
                

                    // 将字段名称首字母大写
                    if (!string.IsNullOrEmpty(fieldName))
                    {
                        fieldName = char.ToUpper(fieldName[0]) + fieldName.Substring(1);
                    }
                    classBuilder.AppendLine($"    public {fieldType} {fieldName} {{ get; set; }}");
                    fieldLoadBuilder.AppendLine($"       this.{fieldName} ={ConvertUtils.GetLoadFieldMethod(fieldType,i)};");
                }
            }

            // 添加获取单个值的方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public static {className} GetById(int id)");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine("        if (_dataDic.TryGetValue(id, out var value))");
            classBuilder.AppendLine("        {");
            classBuilder.AppendLine("            return value;");
            classBuilder.AppendLine("        }");
            classBuilder.AppendLine("        return null;");
            classBuilder.AppendLine("    }");

            // 添加获取值列表的方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public static List<{className}> GetAll()");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine("        if (_dataList == null)");
            classBuilder.AppendLine("        {");
            classBuilder.AppendLine($"            _dataList = new List<{className}>(_dataDic.Values);");
            classBuilder.AppendLine("        }");
            classBuilder.AppendLine("        return _dataList;");
            classBuilder.AppendLine("    }");

           
            //添加Load方法 生成数据
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public void Load(string[] fields)");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine(           fieldLoadBuilder.ToString());
            classBuilder.AppendLine("    }");

            // 添加类结束标记
            classBuilder.AppendLine("}");

            // 将生成的类写入文件
            File.WriteAllText(outputFilePath, classBuilder.ToString(), Encoding.UTF8);
        }

        Console.WriteLine($"C# 类已生成并保存到 {outputFilePath}");
    }

    //读取路径下的所有 Excel 文件，将其转换为 CSV 文件
    public static void ConvertAll(string inputDir, string outputDir)
    {
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var files = Directory.GetFiles(inputDir, "*.xlsx");
        foreach (var file in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(file);
            var outputFilePath = Path.Combine(outputDir, $"{fileName}.cs");
            ConvertToCsharp(file, outputFilePath);
        }
    }
}