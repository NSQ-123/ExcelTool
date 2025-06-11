using System.Text.RegularExpressions;

namespace GameFramework.Table;

public class TableLoaderUtils
{
    private const string CSV_PATTERN = ",(?=(?:[^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)";

    private static string CSV_PATH
    {
        get { return Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../csvOutput")); }
    }

    public static async Task LoadAll<T>(string fileName,Dictionary<int,T> map) where T : ITable, new()
    {
        string csvFilePath = Path.Combine(CSV_PATH, $"{fileName}.csv");
        if (!File.Exists(csvFilePath))
        {
            Console.WriteLine($"CSV file not found: {csvFilePath}");
            return;
        }

        string[] lines = await File.ReadAllLinesAsync(csvFilePath);
        if (lines.Length == 0)
        {
            Console.WriteLine($"CSV file is empty: {csvFilePath}");
            return;
        }

        // 遍历每一行数据
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
            {
                continue; // 跳过空行
            }
            try
            {
                line = line.Trim('\r');
                string[] rowValues = Regex.Split(line, CSV_PATTERN);
                for (int j = 0; j < rowValues.Length; j++)
                {
                    rowValues[j] = rowValues[j].Trim('"');
                }

                // 假设每行数据都可以转换为 className 类型的对象
                // 这里需要根据实际情况进行调整
                T t = new T();
                t.Load(rowValues);
                map[t.GetId()] = t;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing line {i + 1} in {fileName}.csv: {ex.Message}");
                Console.WriteLine($"Line content: {line}");
            }
        }
    }


    
}