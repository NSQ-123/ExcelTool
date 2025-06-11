

using GameFramework.Table;

class Program
{
    private static void ReleasePlay(string[] args)
    {
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: Program <inputExcel> <outputCsv> <outputCsharp>");
            return;
        }

        var inputExcel = args[0];
        var outputCsv = args[1];
        var outputCsharp = args[2];

        // 导出 CSV
        Xlsx2Csv.ConvertAll(inputExcel, outputCsv);

        // 导出 C#
        Xlsx2Csharp.ConvertAll(inputExcel, outputCsharp);
    }


    static void Main(string[] args)
    {

        try
        {
            ReleasePlay(args);
            //_ = DebugPlay();
        }
        catch (Exception ex)
        {
            Console.WriteLine("发生异常：" + ex);
        }
        Console.WriteLine("按任意键退出...");
        Console.ReadKey();

    }

    private static async Task DebugPlay()
    {
        var inputExcel = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../excel"));
        var outputCsv = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../csvOutput"));
        var outputCsharp = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../../csharpOutput"));

        //1.导出csv
        //Xlsx2Csv.ConvertAll(inputExcel, outputCsv);

        //2.导出csharp
        //Xlsx2Csharp.ConvertAll(inputExcel, outputCsharp);

        //3.加载所有表数据
        // await TableDataLoader.LoadAll();
        
        
        // T_Person person = T_Person.GetById(1);
        // Console.WriteLine($"===================");
        // Console.WriteLine($"ID: {person.ID}, Name: {person.Name}, Age: {person.Age}, BornTime: {person.BornTime}, Score: {string.Join(", ", person.Score)}");
        // Console.WriteLine("X1 Data:");
        // foreach (var x1 in person.X1)
        // {
        //     Console.WriteLine($"  X1 Args0: {x1.Args0}, Args1: {x1.Args1}");
        // }
        // Console.WriteLine("X2 Data:");
        // foreach (var x2 in person.X2)
        // {
        //     Console.WriteLine($"  X2 Args0: {x2.Args0}, Args1: {x2.Args1}");
        // }
        // Console.WriteLine("Y1 Data:");
        // foreach (var y1 in person.Y1)
        // {
        //     int index = 0;
        //     Console.WriteLine($"inex = {index}===================");
        //     foreach (var item in y1.Args0)
        //     {
        //         Console.WriteLine($"  Y1 ID: {item}");
        //     }
        // }
    }
    
    
    
    
    

    
    
    
    
}