using System.Collections.Generic;

namespace SwiftExcel.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            const string filePath = "C:\\Development\\ExcelWriter\\test.xlsx";
            var sheetNames = new List<string>
            {
                "sheet 1",
                "sheet 2",
                "custom sheet 3"
            };

            using (var excelWriter = new ExcelWriter(filePath, sheetNames))
            {
                excelWriter.Save();
            }
        }
    }
}