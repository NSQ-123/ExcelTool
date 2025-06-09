
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;

class Program
{
    static void Main(string[] args)
    {

        //Xlsx2Csv.ConvertAll("/Users/ttwj/vs/ExcelTool/ExcelTool/excel", "/Users/ttwj/vs/ExcelTool/ExcelTool/csvOutput");  
        Xlsx2Csharp.ConvertAll("/Users/ttwj/vs/ExcelTool/ExcelTool/excel", "/Users/ttwj/vs/ExcelTool/ExcelTool/csharpOutput");
    }
}