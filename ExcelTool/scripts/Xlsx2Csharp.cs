using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Collections.Generic;
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;
using Org.BouncyCastle.Crypto.Parameters; // 通用接口

public class Xlsx2Csharp
{
    // 统一缩进控制
    private static readonly string indent0 = "";
    private static readonly string indent1 = "\t";
    private static readonly string indent2 = indent1 + "\t";
    private static readonly string indent3 = indent2 + "\t";

    private const string NAME_SPACE = "GameFramework.Table";
    private const string DictionaryName = "_dataMap";
    private const string AsyncOperation = "Task"; // 异步操作的类型名


    //读取路径下的所有 Excel 文件，将其转换为 CSV 文件
    public static void ConvertAll(string inputDir, string outputDir)
    {
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var files = Directory.GetFiles(inputDir, "*.xlsx");
        var list = new List<string>(files.Length);
        foreach (var file in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(file);
            var outputFilePath = Path.Combine(outputDir, $"{fileName}.cs");
            var className = ConvertToCsharp(file, outputFilePath);
            list.Add(className);
        }

        CreateTableDataLoader(outputDir, list);
        Debug($"共转换 {files.Length} 个C# 类文件");
    }


    /// <summary>
    /// 将 Excel 文件的第二个工作表转换为 C# 类定义
    /// </summary>
    /// <param name="excelFilePath">Excel 文件路径</param>
    /// <param name="outputFilePath">生成的 C# 文件路径</param>
    public static string ConvertToCsharp(string excelFilePath, string outputFilePath)
    {
        // 获取文件名称作为类名
        var fileName = Path.GetFileNameWithoutExtension(excelFilePath);
        var className = "T_" + fileName;
        // 打开 Excel 文件
        using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet metaSheet = workbook.GetSheetAt(1); // 获取第二个工作表

            // 创建 StringBuilder 用于生成 C# 类定义
            /************************************************************************************/
            StringBuilder classBuilder = new StringBuilder();
            classBuilder.AppendLine($"{indent0}using System;");
            classBuilder.AppendLine($"{indent0}using System.Collections.Generic;");
            classBuilder.AppendLine();
            /************************************************************************************/
            if (!string.IsNullOrEmpty(NAME_SPACE))
            {
                classBuilder.AppendLine($"{indent0}namespace {NAME_SPACE}");
                classBuilder.AppendLine($"{indent0}{{");
            }

            classBuilder.AppendLine($"{indent1}public partial class {className} : ITable");
            classBuilder.AppendLine($"{indent1}{{");

            // 添加静态字典字段
            classBuilder.AppendLine($"{indent2}private static readonly Dictionary<int, {className}> {DictionaryName} = new Dictionary<int, {className}>();");
            classBuilder.AppendLine($"{indent2}private static List<{className}> _dataList;");
            classBuilder.AppendLine();

            // 遍历字段定义
            IRow fieldNameRow = metaSheet.GetRow(0); // 第一行：字段名称
            IRow fieldTypeRow = metaSheet.GetRow(1); // 第二行：字段类型
            IRow usageRow = metaSheet.GetRow(2); // 第三行：使用方
            IRow descriptionRow = metaSheet.GetRow(3); // 第四行：字段描述

            if (fieldNameRow == null || fieldTypeRow == null || usageRow == null || descriptionRow == null)
            {
                throw new Exception("工作表格式不正确，缺少必要的字段定义行。");
            }

            //处理字段定义
            if (fieldNameRow.LastCellNum != fieldTypeRow.LastCellNum ||
                fieldNameRow.LastCellNum != usageRow.LastCellNum)
            {
                throw new Exception("工作表格式不正确，字段定义行的单元格数量不一致。");
            }

            StringBuilder fieldLoadBuilder = new StringBuilder();
            StringBuilder subClassBuilder = null;

            for (int i = 0; i < fieldNameRow.LastCellNum; i++)
            {
                //处理空格
                if (fieldNameRow.GetCell(i) == null || fieldTypeRow.GetCell(i) == null)
                {
                    continue; // 跳过空单元格
                }

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
                        classBuilder.AppendLine($"{indent2}/// <summary>");
                        classBuilder.AppendLine($"{indent2}/// {description}");
                        classBuilder.AppendLine($"{indent2}/// </summary>");
                    }

                    // 将字段名称首字母大写
                    if (!string.IsNullOrEmpty(fieldName))
                    {
                        fieldName = char.ToUpper(fieldName[0]) + fieldName.Substring(1);
                    }

                    // 添加字段定义
                    fieldType = CorrectType(fieldType); // 规范化字段类型
                    var isArray = fieldType.ToLowerInvariant().StartsWith("arr<") && fieldType.EndsWith(">");
                    var arrType = string.Empty;

                    if (isArray)
                    {
                        arrType = $"T_{fieldName}";
                        subClassBuilder ??= new StringBuilder();
                        ProcessArr(fieldType, arrType, subClassBuilder);

                        classBuilder.AppendLine($"{indent2}public List<{arrType}> {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine($"{indent3}this.{fieldName} = ConvertUtils.LoadArr<{arrType}>(data[{i}]);");
                    }
                    else
                    {
                        classBuilder.AppendLine($"{indent2}public {fieldType} {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine($"{indent3}this.{fieldName} = {GetLoadFieldMethod(fieldType, i)};");
                    }
                }
            }

            // //添加向字典中添加数据的方法
            // classBuilder.AppendLine();
            // classBuilder.AppendLine($"    public static void AddData(int id, {className} data)");
            // classBuilder.AppendLine("    {");
            // classBuilder.AppendLine($"        {DictionaryName}[id] = data;");
            // classBuilder.AppendLine("    }");


            // 添加获取单个值的方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"{indent2}public static {className} GetById(int id)");
            classBuilder.AppendLine($"{indent2}{{");
            classBuilder.AppendLine($"{indent3}if ({DictionaryName}.TryGetValue(id, out var value))");
            classBuilder.AppendLine($"{indent3}{{");
            classBuilder.AppendLine($"{indent3}\treturn value;");
            classBuilder.AppendLine($"{indent3}}}");
            classBuilder.AppendLine($"{indent3}return null;");
            classBuilder.AppendLine($"{indent2}}}");

            // 添加获取值列表的方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"{indent2}public static List<{className}> GetAll()");
            classBuilder.AppendLine($"{indent2}{{");
            classBuilder.AppendLine($"{indent3}if (_dataList == null)");
            classBuilder.AppendLine($"{indent3}{{");
            classBuilder.AppendLine($"{indent3}\t_dataList = new List<{className}>({DictionaryName}.Values);");
            classBuilder.AppendLine($"{indent3}}}");
            classBuilder.AppendLine($"{indent3}return _dataList;");
            classBuilder.AppendLine($"{indent2}}}");


            //添加Load方法 生成数据
            classBuilder.AppendLine();
            classBuilder.AppendLine($"{indent2}public void Load(string[] data)");
            classBuilder.AppendLine($"{indent2}{{");
            classBuilder.Append(fieldLoadBuilder.ToString());
            classBuilder.AppendLine($"{indent2}}}");


            // 添加 GetId 方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"{indent2}public int GetId()");
            classBuilder.AppendLine($"{indent2}{{");
            classBuilder.AppendLine($"{indent3}var idProperty = this.GetType().GetProperty(\"ID\");");
            classBuilder.AppendLine($"{indent3}if (idProperty != null)");
            classBuilder.AppendLine($"{indent3}{{");
            classBuilder.AppendLine($"{indent3}\treturn (int)idProperty.GetValue(this);");
            classBuilder.AppendLine($"{indent3}}}");
            classBuilder.AppendLine($"{indent3}throw new Exception(\"当前类 {{this.GetType().Name}}  未定义 ID 属性\");");
            classBuilder.AppendLine($"{indent2}}}");


            //添加LoadAll方法 加载原始数据
            classBuilder.AppendLine();
            classBuilder.AppendLine($"{indent2}public static async {AsyncOperation} LoadAll(string type)");
            classBuilder.AppendLine($"{indent2}{{");
            classBuilder.AppendLine($"{indent3}await TableLoaderUtils.LoadAll(type, {DictionaryName});");
            classBuilder.AppendLine($"{indent2}}}");


            // 添加类结束标记
            classBuilder.AppendLine($"{indent1}}}");
            /************************************************************************************/

            // 如果存在 arr<...> 类型的字段，则生成对应的子类
            if (subClassBuilder != null && subClassBuilder.Length > 0)
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

        Debug($"生成 C# 类： {className}");

        //return !string.IsNullOrEmpty(NAME_SPACE)?$"{NAME_SPACE}.{className}": className;
        return className; // 返回类名
    }

    private static void ProcessArr(string fieldType, string className, StringBuilder subBuilder)
    {
        subBuilder.AppendLine($"{indent1}public partial class {className} : ITable");
        subBuilder.AppendLine($"{indent1}{{");
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
            subBuilder.AppendLine($"{indent2}public List<{GetCSharpBaseType(baseType)}> Args0;");
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

                subBuilder.AppendLine($"{indent2}public {fieldTypeStr} Args{i};");
            }
        }

        // 实现 ITable 接口

        // 添加 Load 方法
        subBuilder.AppendLine($"{indent2}public void Load(string[] data)");
        subBuilder.AppendLine($"{indent2}{{");
        for (int i = 0; i < typeList.Length; i++)
        {
            string t = typeList[i].Trim();
            // 处理 slice 类型  只支持arr里有一个slice类型
            if (t.EndsWith("slice"))
            {
                string baseType = t.Replace("slice", "").Trim();
                if (string.IsNullOrEmpty(baseType)) baseType = "int";
                var fullType = GetCSharpBaseType(baseType);
                subBuilder.AppendLine($"{indent3}Args{i} = ConvertUtils.GetList<{fullType}>(data);");
                break;
            }
            else
            {
                var fullType = GetCSharpBaseType(t);
                subBuilder.AppendLine($"{indent3}Args{i} = ConvertUtils.Get<{fullType}>(data[{i}]);");
            }
        }

        subBuilder.AppendLine($"{indent2}}}");

        // 添加 GetId 方法
        subBuilder.AppendLine();
        subBuilder.AppendLine($"{indent2}public int GetId()");
        subBuilder.AppendLine($"{indent2}{{");
        subBuilder.AppendLine($"{indent3}var idProperty = this.GetType().GetProperty(\"ID\");");
        subBuilder.AppendLine($"{indent3}if (idProperty != null)");
        subBuilder.AppendLine($"{indent3}{{");
        subBuilder.AppendLine($"{indent3}\treturn (int)idProperty.GetValue(this);");
        subBuilder.AppendLine($"{indent3}}}");
        subBuilder.AppendLine($"{indent3}throw new Exception($\"当前类 {{this.GetType().Name}}  未定义 ID 属性\");");
        subBuilder.AppendLine($"{indent2}}}");

        subBuilder.AppendLine($"{indent1}}}");
        subBuilder.AppendLine();
    }

    private static string GetCSharpBaseType(string type)
    {
        switch (type.ToLowerInvariant())
        {
            case "int": return "System.Int32";
            case "string": return "System.String";
            case "bool": return "System.Boolean";
            case "float": return "System.Single";
            case "double": return "System.Double";
            case "long": return "System.Int64";
            case "datetime": return "System.DateTime";
            default: throw new Exception($"不支持的类型: {type}，请检查字段定义。");
        }
    }

    private static string CorrectType(string fieldType)
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
            throw new Exception("字段类型不能为空");

        fieldType = fieldType.Trim();

        if (fieldType.StartsWith("List<") && fieldType.EndsWith(">"))
        {
            var innerType = fieldType.Substring(5, fieldType.Length - 6).Trim();
            var fullType = GetCSharpBaseType(innerType);

            if (typeof(IConvertible).IsAssignableFrom(Type.GetType(fullType)))
            {
                return $"ConvertUtils.GetList<{fullType}>(data[{index}])";
            }
        }
        else
        {
            var fullType = GetCSharpBaseType(fieldType);
            if (typeof(IConvertible).IsAssignableFrom(Type.GetType(fullType)))
            {
                return $"ConvertUtils.Get<{fullType}>(data[{index}])";
            }
        }

        throw new Exception($"不支持的字段类型: {fieldType}，请检查字段定义。");
    }

    private static void CreateTableDataLoader(string outputDir, List<string> classNames)
    {
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        StringBuilder loaderBuilder = new StringBuilder();
        loaderBuilder.AppendLine($"{indent0}using System;");
        loaderBuilder.AppendLine($"{indent0}using System.Collections.Generic;");
        loaderBuilder.AppendLine();
        if (!string.IsNullOrEmpty(NAME_SPACE))
        {
            loaderBuilder.AppendLine($"{indent0}namespace {NAME_SPACE}");
            loaderBuilder.AppendLine($"{indent0}{{");
        }

        loaderBuilder.AppendLine();
        loaderBuilder.AppendLine($"{indent1}public class TableDataLoader");
        loaderBuilder.AppendLine($"{indent1}{{");
        loaderBuilder.AppendLine($"{indent2}public static async {AsyncOperation} LoadAll()");
        loaderBuilder.AppendLine($"{indent2}{{");
        loaderBuilder.AppendLine($"{indent3}List<{AsyncOperation}> tasks = new();");

        foreach (var className in classNames)
        {
            loaderBuilder.AppendLine($"{indent3}tasks.Add({className}.LoadAll(\"{className.Substring(2)}\"));");
        }

        loaderBuilder.AppendLine($"{indent3}await {AsyncOperation}.WhenAll(tasks);");
        loaderBuilder.AppendLine($"{indent2}}}");
        loaderBuilder.AppendLine();
        loaderBuilder.AppendLine($"{indent1}}}");
        if (!string.IsNullOrEmpty(NAME_SPACE))
        {
            loaderBuilder.AppendLine($"{indent0}}}");
        }

        string outputFilePath = Path.Combine(outputDir, "TableDataLoader.cs");
        File.WriteAllText(outputFilePath, loaderBuilder.ToString(), Encoding.UTF8);
    }


    private static void Debug(string message)
    {
        Console.WriteLine($"[Xlsx2Csharp] {message}");
    }
}