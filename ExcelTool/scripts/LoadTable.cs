/*
 * 导出的CSV文件 每个数据之间用逗号分隔
 * 注意：如果是复合类型，则会用双引号包裹该数据
 * 例如：intSlice="1,2,3";arr<intSlice>="1,2,3;4,5,6;7";arr<int,string,int>="1,abc,3;4,def,5;7,ghi,6"
*/

using System.Text.RegularExpressions;

public class LoadTable
{

    public const string CSV_PATTERN = ",(?=(?:[^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)";
    
    // public static T LoadCsvToObject<T>(string csvFilePath) where T : ITable, new()
    // {
    //     if (!File.Exists(csvFilePath))
    //     {
    //         Console.WriteLine($"[LoadTable] CSV file not found: {csvFilePath}");
    //         return default;
    //     }
    //     try
    //     {
    //         using (var reader = new StreamReader(csvFilePath))
    //         {
    //             string line;
    //             // 跳过表头
    //             if ((line = reader.ReadLine()) == null)
    //                 return default;
    //             while ((line = reader.ReadLine()) != null)
    //             {
    //                 string[] rowValues = Regex.Split(line, CSV_PATTERN);
    //                 for (int i = 0; i < rowValues.Length; i++)
    //                 {
    //                     rowValues[i] = rowValues[i].Trim('"');
    //                 }
    //                 T obj = new T();
    //                 obj.Load(rowValues);
    //                 return obj; // 这里只读取一行，如需全部请用List<T>
    //             }
    //         }
    //     }
    //     catch (Exception e)
    //     {
    //         Console.WriteLine($"[LoadTable] Error reading CSV file: {e.Message}");
    //     }
    //     return default;
    // }
    //
    // public static List<T> LoadCsvToList<T>(string csvFilePath) where T : ITable, new()
    // {
    //     var result = new List<T>();
    //     if (!File.Exists(csvFilePath))
    //     {
    //         Console.WriteLine($"[LoadTable] CSV file not found: {csvFilePath}");
    //         return result;
    //     }
    //     try
    //     {
    //         using (var reader = new StreamReader(csvFilePath))
    //         {
    //             string line;
    //             // 跳过表头
    //             if ((line = reader.ReadLine()) == null)
    //                 return result;
    //             while ((line = reader.ReadLine()) != null)
    //             {
    //                 string[] rowValues = Regex.Split(line, CSV_PATTERN);
    //                 for (int i = 0; i < rowValues.Length; i++)
    //                 {
    //                     rowValues[i] = rowValues[i].Trim('"');
    //                 }
    //                 T obj = new T();
    //                 obj.Load(rowValues);
    //                 result.Add(obj);
    //             }
    //         }
    //     }
    //     catch (Exception e)
    //     {
    //         Console.WriteLine($"[LoadTable] Error reading CSV file: {e.Message}");
    //     }
    //     return result;
    // }
    //
    // public static void LoadAllCsvToClassInstances(string csvDir, string csharpDir)
    // {
    //     if (!Directory.Exists(csvDir) || !Directory.Exists(csharpDir))
    //     {
    //         Console.WriteLine($"[LoadTable] 目录不存在: {csvDir} 或 {csharpDir}");
    //         return;
    //     }
    //     var csvFiles = Directory.GetFiles(csvDir, "*.csv");
    //     foreach (var csvFile in csvFiles)
    //     {
    //         var fileName = Path.GetFileNameWithoutExtension(csvFile);
    //         // 直接拼接类名
    //         var typeName = $"GameFramework.Table.T_{fileName}";
    //         var type = Type.GetType(typeName);
    //         if (type == null)
    //         {
    //             Console.WriteLine($"[LoadTable] 未找到类型: {typeName}");
    //             continue;
    //         }
    //         var method = typeof(LoadTable).GetMethod("LoadCsvToList").MakeGenericMethod(type);
    //         var list = method.Invoke(null, new object[] { csvFile });
    //         Console.WriteLine($"[LoadTable] 加载 {csvFile} 到类型 {typeName}，共{((System.Collections.ICollection)list).Count}条数据");
    //     }
    // }
}