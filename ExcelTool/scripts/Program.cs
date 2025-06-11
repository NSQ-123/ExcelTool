
using NPOI.XSSF.UserModel; // 用于处理 .xlsx 文件
using NPOI.SS.UserModel;   // 通用接口
using System.IO;
using GameFramework.Table;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        var inputCsv = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../excel"));
        var outputCsv = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../csvOutput"));

        var inputCsharp = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../excel"));
        var outputCsharp = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../csharpOutput"));

        // 导出csv
        // Xlsx2Csv.ConvertAll(inputCsv, outputCsv);

        // //导出csharp
        //Xlsx2Csharp.ConvertAll(inputCsharp, outputCsharp);

        // 加载所有表数据
        await TableDataLoader.LoadAll();
        T_Person person = T_Person.GetById(1);
        Console.WriteLine($"===================");
        Console.WriteLine($"ID: {person.ID}, Name: {person.Name}, Age: {person.Age}, BornTime: {person.BornTime}, Score: {string.Join(", ", person.Score)}");
        Console.WriteLine("X1 Data:");
        foreach (var x1 in person.X1)
        {
            Console.WriteLine($"  X1 Args0: {x1.Args0}, Args1: {x1.Args1}");
        }
        Console.WriteLine("X2 Data:");
        foreach (var x2 in person.X2)
        {
            Console.WriteLine($"  X2 Args0: {x2.Args0}, Args1: {x2.Args1}");
        }
        Console.WriteLine("Y1 Data:");
        foreach (var y1 in person.Y1)
        {
            int index = 0;
            Console.WriteLine($"inex = {index}===================");
            foreach (var item in y1.Args0)
            {
                Console.WriteLine($"  Y1 ID: {item}");
            }
           
        }
    }
}