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
         Console.WriteLine($"[Xlsx2Csv] 共导出 {files.Length} 个CSV 文件");
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

            // 先收集所有 usage 含 "c" 的字段索引、类型、名称，顺序与 C# 结构体一致
            IRow fieldNameRow = metaSheet.GetRow(0); // 字段名
            IRow fieldTypeRow = metaSheet.GetRow(1); // 字段类型
            IRow usageRow = metaSheet.GetRow(2);     // 使用方
            if (fieldNameRow == null || fieldTypeRow == null || usageRow == null)
                throw new Exception("元数据表格式不正确，缺少必要的字段定义行。");

            var exportFields = new List<(int Col, string FieldType, string FieldName)>();
            for (int i = 0; i < fieldNameRow.LastCellNum; i++)
            {
                string fieldName = fieldNameRow.GetCell(i)?.StringCellValue ?? "";
                string fieldType = fieldTypeRow.GetCell(i)?.StringCellValue ?? "string";
                string usage = usageRow.GetCell(i)?.StringCellValue ?? "";
                if (string.IsNullOrEmpty(usage)) continue;
                if (usage.ToLowerInvariant().Contains("c"))
                {
                    exportFields.Add((i, fieldType, fieldName));
                }
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
                        //检查第一列是否有数据，如果没有数据则跳过该行
                        ICell firstCell = rowData.GetCell(0);
                        if (firstCell == null || firstCell.CellType == CellType.Blank)
                        {
                            continue; // 如果第一列为空，则跳过该行
                        }

                        // 只导出 usage 含 c 的字段，顺序与 C# 结构体一致
                        StringBuilder csvLine = new StringBuilder();
                        foreach (var (col, fieldType, fieldName) in exportFields)
                        {
                            ICell cellData = rowData.GetCell(col);
                            var type = fieldType.ToLowerInvariant();
                            string cellStr = cellData != null ? GetCellString(cellData) : string.Empty;
                            // 类型映射与 C# 结构体一致，支持 arr<...>、slice、基础类型
                            csvLine.Append(ProcessCellValue(type, cellStr));
                            csvLine.Append(",");
                        }
                        //移除csvLine末尾的逗号
                        if (csvLine.Length > 0 && csvLine[^1] == ',')
                        {
                            csvLine.Length--;
                        }
                        // 写入 CSV 文件
                        writer.WriteLine(csvLine.ToString());
                    }
                }
            }
        }

        Console.WriteLine($"[Xlsx2Csv] 导出 CSV 文件: {csvFilePath}");
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
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? ((DateTime)cell.DateCellValue).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            _ => cell.ToString()
        };
    }

    private static string ProcessCellValue(string type, string cellStr)
    {
        // arr<> 结构，直接返回原有字符串（如需更复杂的结构可扩展）
        if (type.StartsWith("arr<") && type.EndsWith(">"))
            return ProcessArr(type, cellStr, false);

        // slice 类型统一用 ConvertUtils 泛型
        switch (type)
        {
            case "intslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<int>(cellStr)) + '"';
            case "boolslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<bool>(cellStr)) + '"';
            case "floatslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<float>(cellStr)) + '"';
            case "doubleslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<double>(cellStr)) + '"';
            case "stringslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<string>(cellStr)) + '"';
            case "longslice":
                return '"' + string.Join(",", GameFramework.Table.ConvertUtils.GetList<long>(cellStr)) + '"';
            case "datetimeslice":
                var list = GameFramework.Table.ConvertUtils.GetList<System.DateTime>(cellStr);
                return '"' + string.Join(",", list.ConvertAll(dt => dt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture))) + '"';
        }

        // 基础类型统一用 ConvertUtils 泛型
        switch (type)
        {
            case "int":
                return GameFramework.Table.ConvertUtils.Get<int>(cellStr).ToString();
            case "float":
                return GameFramework.Table.ConvertUtils.Get<float>(cellStr).ToString();
            case "double":
                return GameFramework.Table.ConvertUtils.Get<double>(cellStr).ToString();
            case "long":
                return GameFramework.Table.ConvertUtils.Get<long>(cellStr).ToString();
            case "string":
                return cellStr;
            case "bool":
                return GameFramework.Table.ConvertUtils.Get<bool>(cellStr).ToString();
            case "datetime":
                if (string.IsNullOrEmpty(cellStr)) return string.Empty;
                var dt = GameFramework.Table.ConvertUtils.Get<System.DateTime>(cellStr);
                return dt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        }
        return string.Empty;
    }

    // 旧的类型转换方法已被统一的泛型方法替代
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
            var groupContent = trimmedGroup.Trim('(', ')');
            if (typeList.Length == 1 && typeList[0].EndsWith("slice"))
            {
                // arr<intSlice> 特殊处理：组内全是int，用逗号分割，不包裹引号
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
                    // arr<> 里嵌套 slice 也不包裹引号
                    itemResult.Add(ProcessArrItem(t, v, false));
                }
                var groupStr = string.Join(",", itemResult);
                result.Add(useGroupParentheses ? $"({groupStr})" : groupStr);
            }
        }
        return '"' + string.Join(";", result) + '"';
    }

    // 处理 arr<> 内部每个元素的类型
    private static string ProcessArrItem(string type, string value, bool wrapQuote = false)
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
            default:
                // slice类型嵌套，直接用ProcessSlice，arr<> 里 wrapQuote=false，单独slice字段 wrapQuote=true
                if (type.EndsWith("slice"))
                    return ProcessSlice(type, value, wrapQuote);
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
        return wrapQuote ? '"' + joined + '"' : joined;
    }
}
