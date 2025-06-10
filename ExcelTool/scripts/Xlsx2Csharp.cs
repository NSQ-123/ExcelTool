using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口

public class Xlsx2Csharp
{
    private const string NAME_SPACE = "GameFramework.Table";
    private const string DictionaryName = "DataMap";
    
    
    
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
            /************************************************************************************/
            StringBuilder classBuilder = new StringBuilder();
            classBuilder.AppendLine("using System;");
            classBuilder.AppendLine("using System.Collections.Generic;");
            classBuilder.AppendLine();
            /************************************************************************************/
            if (!string.IsNullOrEmpty(NAME_SPACE))
            {
                classBuilder.AppendLine($"namespace {NAME_SPACE}");
                classBuilder.AppendLine("{");
            }
            classBuilder.AppendLine($"public partial class {className} : ITable");
            classBuilder.AppendLine("{");

            // 添加静态字典字段
            classBuilder.AppendLine($"    public static Dictionary<int, {className}> {DictionaryName} = new Dictionary<int, {className}>();");
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
            StringBuilder subClassBuilder = null;
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

                    // 将字段名称首字母大写
                    if (!string.IsNullOrEmpty(fieldName))
                    {
                        fieldName = char.ToUpper(fieldName[0]) + fieldName.Substring(1);
                    }
                    
                    // 添加字段定义
                    fieldType = GetCsharpType(fieldType); // 规范化字段类型
                    var isArray = fieldType.ToLowerInvariant().StartsWith("arr<") && fieldType.EndsWith(">");
                    var arrType = string.Empty;
                    
                    if (isArray)
                    {
                        arrType = $"T_{fieldName}";
                        subClassBuilder??= new StringBuilder();
                        ProcessArr(fieldType,arrType,subClassBuilder);
                    }

                  
                    if (!isArray)
                    {
                        classBuilder.AppendLine($"    public {fieldType} {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine($"        this.{fieldName} = {GetLoadFieldMethod(fieldType, i)};");
                    }
                    else
                    {
                        classBuilder.AppendLine($"    public List<{arrType}> {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine($"        this.{fieldName} = ConvertUtils.LoadArgs<{arrType}>(data[{i}]);");
                    }
                    
                }
            }

            // 添加获取单个值的方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public static {className} GetById(int id)");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine($"       if ({DictionaryName}.TryGetValue(id, out var value))");
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
            classBuilder.AppendLine($"            _dataList = new List<{className}>({DictionaryName}.Values);");
            classBuilder.AppendLine("        }");
            classBuilder.AppendLine("        return _dataList;");
            classBuilder.AppendLine("    }");

           
            //添加Load方法 生成数据
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public void Load(string[] data)");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine(           fieldLoadBuilder.ToString());
            classBuilder.AppendLine("    }");

            // 添加类结束标记
            classBuilder.AppendLine("}");
            /************************************************************************************/
            
            // 如果存在 arr<...> 类型的字段，则生成对应的子类
            if(subClassBuilder != null && subClassBuilder.Length > 0)
            {
                classBuilder.AppendLine();
                classBuilder.Append(subClassBuilder.ToString());
            }
            
            // 如果指定了命名空间，则添加命名空间结束标记
            if (!string.IsNullOrEmpty(NAME_SPACE))
            {
                classBuilder.AppendLine("}");
            }
            /************************************************************************************/
            // 将生成的类写入文件
            File.WriteAllText(outputFilePath, classBuilder.ToString(), Encoding.UTF8);
        }

        Console.WriteLine($"C# 类已生成并保存到 {outputFilePath}");
    }

    private static void ProcessArr(string fieldType,string className,StringBuilder subBuilder)
    {
        subBuilder.AppendLine($"public partial class {className} : ITable");
        subBuilder.AppendLine("{");
        /*******************************************************************************************/
        // 处理 arr<...> 类型的字段
        var innerType = fieldType.Substring(4, fieldType.Length - 5).ToLowerInvariant();
        var typeList = innerType.Split(',');
        if (typeList.Length == 1 && typeList[0].EndsWith("slice"))
        {
            // 只有一个slice类型，生成一个List<基础类型>字段
            string baseType = typeList[0].Replace("slice", "").Trim();
            if (string.IsNullOrEmpty(baseType)) baseType = "int";
            baseType = baseType.ToLowerInvariant();
            subBuilder.AppendLine($"    public List<{GetCSharpBaseType(baseType)}> Args0;");
        }
        else
        {
            // 多个类型，生成 Args0, Args1, ...
            for (int i = 0; i < typeList.Length; i++)
            {
                string t = typeList[i].Trim();
                string fieldTypeStr;
                if (t.EndsWith("slice"))
                {
                    string baseType = t.Replace("slice", "").Trim();
                    if (string.IsNullOrEmpty(baseType)) baseType = "int";
                    fieldTypeStr = $"List<{GetCSharpBaseType(baseType)}>";
                }
                else
                {
                    fieldTypeStr = GetCSharpBaseType(t);
                }
                subBuilder.AppendLine($"    public {fieldTypeStr} Args{i};");
            }
        }
        
        // 实现 ITable 接口

        // 添加 Load 方法
        subBuilder.AppendLine($"    public void Load(string[] data)");
        subBuilder.AppendLine("    {");
        for (int i = 0; i < typeList.Length; i++)
        {
            string t = typeList[i].Trim();
            if (t.EndsWith("slice"))
            {
                string baseType = t.Replace("slice", "").Trim();
                if (string.IsNullOrEmpty(baseType)) baseType = "int";
                var methodName = GetCSharpBaseType(baseType);
                //methodName 首字母大写
                methodName = char.ToUpper(methodName[0]) + methodName.Substring(1);
                subBuilder.AppendLine($"        Args{i} = ConvertUtils.Get{methodName}List(data[{i}]);");
            }
            else
            {
                var methodName = GetCSharpBaseType(t);
                //methodName 首字母大写
                methodName = char.ToUpper(methodName[0]) + methodName.Substring(1);
                subBuilder.AppendLine($"        Args{i} = ConvertUtils.Get{methodName}(data[{i}]);");
            }
        }
        subBuilder.AppendLine("    }");


        subBuilder.AppendLine("}");
        subBuilder.AppendLine();
    }

    private static string GetCSharpBaseType(string type)
    {
        switch (type.ToLowerInvariant())
        {
            case "int": return "int";
            case "string": return "string";
            case "bool": return "bool";
            case "float": return "float";
            case "double": return "double";
            case "long": return "long";
            case "datetime": return "DateTime";
            default: return "string";
        }
    }


    private static string GetCsharpType(string fieldType)
    {
        if (string.IsNullOrEmpty(fieldType))
        {
            return "string"; // 默认类型为 string
        }

        return fieldType.ToLowerInvariant() switch
        {
            "int" => "int",
            "float" => "float",
            "double" => "double",
            "string" => "string",
            "bool" => "bool",
            "long" => "long",
            "datetime" => "DateTime",
            "intslice" => "List<int>",
            "boolslice" => "List<bool>",
            "floatslice" => "List<float>",
            "doubleslice" => "List<double>",
            "stringslice" => "List<string>",
            "longslice" => "List<long>",
            "datetimeslice" => "List<DateTime>",
            _ => fieldType
        };
    }

    private static string GetLoadFieldMethod(string fieldType, int index)
    {
        if (string.IsNullOrEmpty(fieldType))
        {
            return $"ConvertUtils.GetString(data[{index}])"; // 默认类型为 string
        }

        switch (fieldType)
        {
            case "int":
                return $"ConvertUtils.GetInt(data[{index}])";
            case "float":
                return $"ConvertUtils.GetFloat(data[{index}])";
            case "double":
                return $"ConvertUtils.GetDouble(data[{index}])";
            case "string":
                return $"ConvertUtils.GetString(data[{index}])";
            case "bool":
                return $"ConvertUtils.GetBool(data[{index}])";
            case "long":
                return $"ConvertUtils.GetLong(data[{index}])";
            case "DateTime":
                return $"ConvertUtils.GetDateTime(data[{index}])";
            case "List<int>":
                return $"ConvertUtils.GetIntList(data[{index}])";
            case "List<bool>":
                return $"ConvertUtils.GetBoolList(data[{index}])";
            case "List<float>":
                return $"ConvertUtils.GetFloatList(data[{index}])";
            case "List<double>":
                return $"ConvertUtils.GetDoubleList(data[{index}])";
            case "List<string>":
                return $"ConvertUtils.GetStringList(data[{index}])";
            case "List<long>":
                return $"ConvertUtils.GetLongList(data[{index}])";
            case "List<DateTime>":
                return $"ConvertUtils.GetDateTimeList(data[{index}])";
            default:
                throw new NotSupportedException($"不支持的字段类型: {fieldType}");
        }
    }
    
    
}
