
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // 导出csv
        // var inputCsv = Path.Combine(Directory.GetCurrentDirectory(), "../../../excel");
        // var outputCsv = Path.Combine(Directory.GetCurrentDirectory(), "../../../csvOutput");
        // Xlsx2Csv.ConvertAll(inputCsv, outputCsv);


        // //导出csharp
        var inputCsharp = Path.Combine(Directory.GetCurrentDirectory(), "../../../excel");
        var outputCsharp = Path.Combine(Directory.GetCurrentDirectory(), "../../../csharpOutput");
        Xlsx2Csharp.ConvertAll(inputCsharp, outputCsharp);
        
    }
}