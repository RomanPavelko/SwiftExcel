using System.Collections.Generic;

namespace SwiftExcel.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            const string filePath = "C:\\Development\\ExcelWriter\\test.xlsx";
            var sheets = new List<Sheet>
            {
                new Sheet { Name = "sheet 1", ColumnsWidth = new List<double> { 10, 12.12, 30, 25.1, 8, 20 } },
                new Sheet { Name = "sheet 2", ColumnsWidth = new List<double> { 5, 4, 12, 10.5, 9.3, 27.12 } },
                new Sheet { Name = "custom sheet 3", ColumnsWidth = new List<double> { 9.8, 5, 6, 10, 10, 13 } }
            };

            using (var excelWriter = new ExcelWriter(filePath, sheets))
            {
                excelWriter.Save();
            }
        }
    }
}