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
                new Sheet { Name = "custom sheet 3" }
            };

            using (var excelWriter = new ExcelWriter(filePath, sheets))
            {
                excelWriter.Write("start", 1, 1, 1);
                excelWriter.Write("102.25", 2, 3, 1, DataType.Number);
                excelWriter.Write("10/12/2012", 3, 3, 1);
                excelWriter.Write("column 7 row 3", 7, 3, 1);
                excelWriter.Write("column 8 row 4", 8, 4, 1);
                excelWriter.Write("column 2 row 7", 2, 7, 1);
                excelWriter.Write("column 5 row 10", 5, 10, 1);

                excelWriter.Write("second sheet", 2, 8, 2);

                excelWriter.Write("3rd sheet", 3, 3, 3);
                excelWriter.Write("3rd sheet", 800, 800, 3);
                
                excelWriter.Save();
            }
        }
    }
}