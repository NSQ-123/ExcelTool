using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;
using System.Text;
using System.Globalization;


public class Xlsx2Csv
{

    public static void Convert(string xlsxFilePath, string csvFilePath)
    {
        // 打开 Excel 文件
        using (FileStream fileStream = new FileStream(xlsxFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0); // 获取第一个工作表

            // 创建 CSV 文件
            using (StreamWriter writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
            {
                // 遍历工作表中的每一行
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    //第1行是字段名称 跳过
                    if (i == 0) continue;
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // 遍历行中的每个单元格
                        StringBuilder csvLine = new StringBuilder();
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                // 根据单元格类型获取值
                                switch (cell.CellType)
                                {
                                    case CellType.String:
                                        csvLine.Append(cell.StringCellValue);
                                        break;
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell)) // 检查是否为日期
                                        {
                                            if (cell.DateCellValue != null)
                                            {
                                                // 如果是日期类型，使用 ToString 格式化
                                                // 注意：DateCellValue 返回的是 DateTime 类型
                                                var date = (DateTime)cell.DateCellValue;
                                                var dateStr = date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                                                csvLine.Append(dateStr);
                                            }
                                            else
                                            {
                                                // 如果 DateCellValue 为 null，可能是因为单元格为空或格式不正确
                                                csvLine.Append(""); // 写入空字符串
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            csvLine.Append(cell.NumericCellValue);
                                        }
                                        break;
                                    case CellType.Boolean:
                                        csvLine.Append(cell.BooleanCellValue);
                                        break;
                                    default:
                                        csvLine.Append("");
                                        break;
                                }
                            }
                            csvLine.Append(","); // 添加逗号分隔符
                        }

                        // 写入 CSV 文件
                        writer.WriteLine(csvLine.ToString().TrimEnd(','));
                    }
                }
            }
        }

        Console.WriteLine($"CSV 文件已导出到 {csvFilePath}");
    }


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


} 
