using EVM_ECLR.Classes;

class Program
{
    static void Main(string[] args)
    {
        ExcelGenerator.CreateExcelFile(
            filePath: "Merged.xlsx",
            jitPath: "T_JIT.xlsx",
            forecastPath: "Forecast.xlsx",
            totalHoursPath: "Total_Hours.xlsx",
            routingPath: "Routing_PL.xlsx",
            backlogPath: "BACKLOG_Export.xlsx"
        );

        Console.WriteLine("Merge completed!");
    }
}

