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
        Debug($"已将 {files.Length} 个 Excel 文件转换为 C# 类定义，并保存到 {outputDir}");
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
            classBuilder.AppendLine(
                $"    private static readonly Dictionary<int, {className}> {DictionaryName} = new Dictionary<int, {className}>();");
            classBuilder.AppendLine($"    private static List<{className}> _dataList;");
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
                    fieldType = CorrectType(fieldType); // 规范化字段类型
                    var isArray = fieldType.ToLowerInvariant().StartsWith("arr<") && fieldType.EndsWith(">");
                    var arrType = string.Empty;

                    if (isArray)
                    {
                        arrType = $"T_{fieldName}";
                        subClassBuilder ??= new StringBuilder();
                        ProcessArr(fieldType, arrType, subClassBuilder);

                        classBuilder.AppendLine($"    public List<{arrType}> {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine(
                            $"        this.{fieldName} = ConvertUtils.LoadArr<{arrType}>(data[{i}]);");
                    }
                    else
                    {
                        classBuilder.AppendLine($"    public {fieldType} {fieldName} {{ get; set; }}");
                        fieldLoadBuilder.AppendLine($"        this.{fieldName} = {GetLoadFieldMethod(fieldType, i)};");
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
            classBuilder.AppendLine(fieldLoadBuilder.ToString());
            classBuilder.AppendLine("    }");


            // 添加 GetId 方法
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public int GetId()");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine($"        var idProperty = this.GetType().GetProperty(\"ID\");");
            classBuilder.AppendLine($"        if (idProperty != null)");
            classBuilder.AppendLine("        {");
            classBuilder.AppendLine("            return (int)idProperty.GetValue(this);");
            classBuilder.AppendLine("        }");
            classBuilder.AppendLine("        throw new Exception($\"当前类 {this.GetType().Name} 未定义 ID 属性\");");
            classBuilder.AppendLine("    }");


            //添加LoadAll方法 加载原始数据
            classBuilder.AppendLine();
            classBuilder.AppendLine($"    public static async {AsyncOperation} LoadAll(string type)");
            classBuilder.AppendLine("    {");
            classBuilder.AppendLine($"       await TableLoaderUtils.LoadAll(type, {DictionaryName});");
            classBuilder.AppendLine("    }");


            // 添加类结束标记
            classBuilder.AppendLine("}");
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

        Debug($"C# 类已生成并保存到 {outputFilePath}");

        //return !string.IsNullOrEmpty(NAME_SPACE)?$"{NAME_SPACE}.{className}": className;
        return className; // 返回类名
    }

    private static void ProcessArr(string fieldType, string className, StringBuilder subBuilder)
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
            // 处理 slice 类型  只支持arr里有一个slice类型
            if (t.EndsWith("slice"))
            {
                string baseType = t.Replace("slice", "").Trim();
                if (string.IsNullOrEmpty(baseType)) baseType = "int";
                var fullType = GetCSharpBaseType(baseType);
                subBuilder.AppendLine($"        Args{i} = ConvertUtils.GetList<{fullType}>(data);");
                break;
            }
            else
            {
                var fullType = GetCSharpBaseType(t);
                subBuilder.AppendLine($"        Args{i} = ConvertUtils.Get<{fullType}>(data[{i}]);");
            }
        }

        subBuilder.AppendLine("    }");

        // 添加 GetId 方法
        subBuilder.AppendLine();
        subBuilder.AppendLine($"    public int GetId()");
        subBuilder.AppendLine("    {");
        subBuilder.AppendLine($"        var idProperty = this.GetType().GetProperty(\"ID\");");
        subBuilder.AppendLine($"        if (idProperty != null)");
        subBuilder.AppendLine("        {");
        subBuilder.AppendLine("            return (int)idProperty.GetValue(this);");
        subBuilder.AppendLine("        }");
        subBuilder.AppendLine("        throw new Exception($\"当前类 {this.GetType().Name} 未定义 ID 属性\");");
        subBuilder.AppendLine("    }");


        subBuilder.AppendLine("}");
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
        loaderBuilder.AppendLine("using System;");
        loaderBuilder.AppendLine("using System.Collections.Generic;");
        if (!string.IsNullOrEmpty(NAME_SPACE))
        {
            loaderBuilder.AppendLine($"namespace {NAME_SPACE}");
            loaderBuilder.AppendLine("{");
        }

        loaderBuilder.AppendLine();
        loaderBuilder.AppendLine("public class TableDataLoader");
        loaderBuilder.AppendLine("{");

        loaderBuilder.AppendLine($"    public static async {AsyncOperation} LoadAll()");
        loaderBuilder.AppendLine("    {");
        loaderBuilder.AppendLine($"      List<{AsyncOperation}> tasks = new();");

        foreach (var className in classNames)
        {
            loaderBuilder.AppendLine($"      tasks.Add({className}.LoadAll(\"{className.Substring(2)}\"));");
        }

        loaderBuilder.AppendLine($"      await {AsyncOperation}.WhenAll(tasks);");
        loaderBuilder.AppendLine("    }");
        loaderBuilder.AppendLine();
        loaderBuilder.AppendLine("}");
        if (!string.IsNullOrEmpty(NAME_SPACE))
        {
            loaderBuilder.AppendLine("}");
        }

        string outputFilePath = Path.Combine(outputDir, "TableDataLoader.cs");
        File.WriteAllText(outputFilePath, loaderBuilder.ToString(), Encoding.UTF8);
    }


    private static void Debug(string message)
    {
        Console.WriteLine($"[Xlsx2Csharp] {message}");
    }
}