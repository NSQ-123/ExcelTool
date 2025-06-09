
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;

class Program
{
    static void Main(string[] args)
    {

        //Xlsx2Csv.ConvertAll("/Users/ttwj/vs/ExcelTool/ExcelTool/excel", "/Users/ttwj/vs/ExcelTool/ExcelTool/csvOutput");  
        //Xlsx2Csharp.ConvertAll("/Users/ttwj/vs/ExcelTool/ExcelTool/excel", "/Users/ttwj/vs/ExcelTool/ExcelTool/csharpOutput");
        //打印当前文件夹
        var input = Path.Combine(Directory.GetCurrentDirectory(), "../../../excel");
        var output = Path.Combine(Directory.GetCurrentDirectory(), "../../../csharpOutput");
        Console.WriteLine("当前文件夹: " + Directory.GetCurrentDirectory());

        Console.WriteLine(input);
        Console.WriteLine(output);
        Xlsx2Csharp.ConvertAll(input, output);
        
    }
}