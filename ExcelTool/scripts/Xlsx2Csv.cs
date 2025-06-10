using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;
using System.Text;
using System.Globalization;


public class Xlsx2Csv
{
    //读取路径下的所有 Excel 文件，将其转换为 CSV 文件
    public static void ConvertAll(string inputDirectory, string outputDirectory)
    {
        if (!Directory.Exists(inputDirectory))
        {
            throw new DirectoryNotFoundException($"输入目录 '{inputDirectory}' 不存在。");
        }

        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory); // 确保输出目录存在
        }

        var files = Directory.GetFiles(inputDirectory, "*.xlsx");
        foreach (var file in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(file);
            var csvFilePath = Path.Combine(outputDirectory, $"{fileName}.csv");
            Convert(file, csvFilePath);
        }
    }
    
    public static void Convert(string xlsxFilePath, string csvFilePath)
    {
        // 打开 Excel 文件
        using (FileStream fileStream = new FileStream(xlsxFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet dataSheet = workbook.GetSheetAt(0); // 获取第一个工作表
            ISheet metaSheet = workbook.GetSheetAt(1); // 获取第二个工作表
            if (dataSheet == null)
            {
                throw new Exception($"工作表 '{dataSheet.SheetName}' 不存在。");
            }
            if (metaSheet == null)
            {
                throw new Exception($"工作表 '{metaSheet.SheetName}' 不存在。");
            }

            // 创建 CSV 文件
            using (StreamWriter writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
            {
                // 遍历工作表中的每一行
                for (int i = 0; i <= dataSheet.LastRowNum; i++)
                {
                    //第1行是字段名称 跳过
                    if (i == 0) continue;
                    IRow rowData = dataSheet.GetRow(i);
                    if (rowData != null)
                    { 
                        /**********************************************************************/
                        //检查第一列是否有数据，如果没有数据则跳过该行
                        ICell firstCell = rowData.GetCell(0);
                        if (firstCell == null || firstCell.CellType == CellType.Blank)
                        {
                            continue; // 如果第一列为空，则跳过该行
                        }
                        /**********************************************************************/
                        
                        // 遍历行中的每个单元格
                        StringBuilder csvLine = new StringBuilder();
                        for (int j = 0; j < rowData.LastCellNum; j++)
                        {
                            ICell cellData = rowData.GetCell(j);
                            if (cellData != null)
                            {
                                var fieldTuple = GetFieldTuple(metaSheet, j);
                                if(string.IsNullOrEmpty(fieldTuple.FiledType))
                                    continue; // 如果字段类型为空，则跳过该列
                                if(string.IsNullOrEmpty(fieldTuple.FieldUsage))
                                    continue; // 如果使用方为空，则跳过该列
                                if (fieldTuple.FieldUsage.Contains("n", StringComparison.InvariantCultureIgnoreCase))
                                    continue; // 如果使用方为 "n"，则跳过该列
                                /**********************************************************************/
                                // TODO:仅处理客户端使用的字段（含有 "c"）
                                /**********************************************************************/
                                var type = fieldTuple.FiledType.ToLowerInvariant();
                                string cellStr = GetCellString(cellData);
                                csvLine.Append(ProcessCellValue(type, cellStr));
                                csvLine.Append(",");
                            }
                        }
                        
                        //移除csvLine末尾的逗号
                        if (csvLine.Length > 0 && csvLine[^1] == ',')
                        {
                            csvLine.Length--;
                        }
                        // 写入 CSV 文件
                        writer.WriteLine(csvLine.ToString());
                        //writer.WriteLine(csvLine.ToString().TrimEnd(','));
                    }
                }
            }
        }

        Console.WriteLine($"CSV 文件已导出到 {csvFilePath}");
    }



    //获取没列数据的定义类型
    private static (string FiledType, string FieldName, string FieldUsage) GetFieldTuple(ISheet metaSheet, int index)
    {
        // 假设元数据表的第一行是字段名称，第二行是字段类型
        IRow fieldNameRow = metaSheet.GetRow(0);
        IRow fieldTypeRow = metaSheet.GetRow(1);
        IRow fieldUsageRow = metaSheet.GetRow(2); // 第三行：使用方
        if (fieldNameRow == null || fieldTypeRow == null || fieldUsageRow == null)
        {
            throw new Exception("元数据表格式不正确，缺少必要的字段定义行。");
        }

        var filedType = fieldTypeRow.GetCell(index)?.StringCellValue ?? "string"; // 默认类型为 string
        var fieldName = fieldNameRow.GetCell(index)?.StringCellValue ?? "unknown"; // 默认字段名为 unknown
        var fieldUsage = fieldUsageRow.GetCell(index)?.StringCellValue ?? ""; // 获取使用方信息
        return (filedType, fieldName,fieldUsage);
    }






    // 新增统一获取单元格字符串的方法
    private static string GetCellString(ICell cell)
    {
        if (cell == null) return string.Empty;
        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Formula => cell.ToString(),
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? ((DateTime)cell.DateCellValue).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            _ => cell.ToString()
        };
    }

    private static string ProcessCellValue(string type, string cellStr)
    {
        if (type.StartsWith("arr<") && type.EndsWith(">"))
        {
            // 默认不包裹组
            return ProcessArr(type, cellStr, false);
        }
        //intSlice, boolSlice 等
        if (type.EndsWith("slice"))
        {
            return ProcessSlice(type, cellStr, true);
        }
        // 处理基本类型
        switch (type)
        {
            case "int":
                return ProcessInt(cellStr);
            case "float":
                return ProcessFloat(cellStr);
            case "double":
                return ProcessDouble(cellStr);
            case "long":
                return ProcessLong(cellStr);
            case "string":
                return ProcessString(cellStr);
            case "bool":
                return ProcessBool(cellStr);
            case "datetime":
                return ProcessDateTime(cellStr);
            default:
                return string.Empty;
        }
    }

    private static string ProcessInt(string cellStr)
    {
        return int.TryParse(cellStr, out int intValue) ? intValue.ToString() : string.Empty;
    }
    private static string ProcessFloat(string cellStr)
    {
        return float.TryParse(cellStr, out float floatValue) ? floatValue.ToString() : string.Empty;
    }
    private static string ProcessDouble(string cellStr)
    {
        return double.TryParse(cellStr, out double doubleValue) ? doubleValue.ToString() : string.Empty;
    }
    private static string ProcessLong(string cellStr)
    {
        return long.TryParse(cellStr, out long longValue) ? longValue.ToString() : string.Empty;
    }
    private static string ProcessString(string cellStr)
    {
        return cellStr;
    }
    private static string ProcessBool(string cellStr)
    {
        return bool.TryParse(cellStr, out bool boolValue) ? boolValue.ToString() : string.Empty;
    }
    private static string ProcessDateTime(string cellStr)
    {
        return DateTime.TryParse(cellStr, out DateTime dateValue) ? dateValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty;
    }
    // 
    /// <summary>
    /// arr处理方法
    /// </summary>
    /// <param name="type"></param>
    /// <param name="cellStr"></param>
    /// <param name="useGroupParentheses">使用组括号</param>
    /// <returns></returns>
    /// 新增 arr<> 处理方法，支持任意类型组合，组和组之间用;分割，组内用逗号分割，整个数据用双引号包裹，每组可用()包裹
    private static string ProcessArr(string type, string cellStr, bool useGroupParentheses = false)
    {
        if (string.IsNullOrWhiteSpace(cellStr)) return string.Empty;
        var innerType = type.Substring(4, type.Length - 5).ToLowerInvariant();
        var typeList = innerType.Split(',');
        var groups = cellStr.Split(';'); // 组之间用;分割
        var result = new List<string>();
        foreach (var group in groups)
        {
            var trimmedGroup = group.Trim();
            if (string.IsNullOrEmpty(trimmedGroup)) continue;
            var groupContent = trimmedGroup.Trim('(', ')'); // 始终去除外部括号  因为填写表的时候带有括号
            // arr<intSlice> 特殊处理：组内全是int，用逗号分割
            if (typeList.Length == 1 && typeList[0].EndsWith("slice"))
            {
                var sliceStr = ProcessSlice(typeList[0], groupContent, false);
                result.Add(useGroupParentheses ? $"({sliceStr})" : sliceStr);
            }
            else
            {
                var items = groupContent.Split(',');
                var itemResult = new List<string>();
                for (int i = 0; i < typeList.Length && i < items.Length; i++)
                {
                    var t = typeList[i].Trim();
                    var v = items[i].Trim();
                    if (t.EndsWith("slice"))
                        itemResult.Add(ProcessSlice(t, v, false));
                    else
                        itemResult.Add(ProcessArrItem(t, v));
                }
                var groupStr = string.Join(",", itemResult);
                result.Add(useGroupParentheses ? $"({groupStr})" : groupStr);
            }
        }
        return '"' + string.Join(";", result) + '"';
    }

    // 处理 arr<> 内部每个元素的类型
    private static string ProcessArrItem(string type, string value)
    {
        switch (type)
        {
            case "int":
                return int.TryParse(value, out int intValue) ? intValue.ToString() : "";
            case "float":
                return float.TryParse(value, out float floatValue) ? floatValue.ToString() : "";
            case "double":
                return double.TryParse(value, out double doubleValue) ? doubleValue.ToString() : "";
            case "long":
                return long.TryParse(value, out long longValue) ? longValue.ToString() : "";
            case "bool":
                return bool.TryParse(value, out bool boolValue) ? boolValue.ToString() : "";
            case "datetime":
                return DateTime.TryParse(value, out DateTime dateValue) ? dateValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";
            case "string":
                return value;
            case var t when t.EndsWith("slice"):
                // slice类型嵌套，直接用ProcessSlice
                return ProcessSlice(type, value);
            default:
                return value;
        }
    }

    // 支持字符串直接处理slice，增加参数 control 是否包裹引号
    private static string ProcessSlice(string type, string cellStr, bool wrapQuote = true)
    {
        if (string.IsNullOrWhiteSpace(cellStr)) return string.Empty;
        var sliceData = cellStr.Split(',');
        var result = new List<string>();
        foreach (var item in sliceData)
        {
            var trimmed = item.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;
            switch (type)
            {
                case "intslice":
                    if (int.TryParse(trimmed, out int intValue)) result.Add(intValue.ToString());
                    break;
                case "boolslice":
                    if (bool.TryParse(trimmed, out bool boolValue)) result.Add(boolValue.ToString());
                    break;
                case "floatslice":
                    if (float.TryParse(trimmed, out float floatValue)) result.Add(floatValue.ToString());
                    break;
                case "doubleslice":
                    if (double.TryParse(trimmed, out double doubleValue)) result.Add(doubleValue.ToString());
                    break;
                case "stringslice":
                    result.Add(trimmed);
                    break;
                case "longslice":
                    if (long.TryParse(trimmed, out long longValue)) result.Add(longValue.ToString());
                    break;
                case "datetimeslice":
                    if (DateTime.TryParse(trimmed, out DateTime dateTimeValue)) result.Add(dateTimeValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                    break;
                default:
                    result.Add(trimmed);
                    break;
            }
        }
        var joined = string.Join(",", result);
        return wrapQuote ? ('"' + joined + '"') : joined;
    }
}
