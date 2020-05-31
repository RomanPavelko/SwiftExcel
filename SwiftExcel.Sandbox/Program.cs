using SwiftExcel.Extensions;
using System.Collections.Generic;

namespace SwiftExcel.Sandbox
{
    internal class Program
    {
        private const string FilePath = "C:/temp/test.xlsx";

        private static ExcelWriter _excelWriter;

        private static void Main()
        {
            //Fill excel document with test data 100 rows x 10 columns
            using (_excelWriter = new ExcelWriter(FilePath))
            {
                for (var row = 1; row <= 100; row++)
                {
                    for (var col = 1; col <= 10; col++)
                    {
                        _excelWriter.Write($"row:{row}-col:{col}", col, row);
                    }
                }
            }


            //Set custom sheet name and define columns width
            //Use manual Save() instead of using block 
            var sheet = new Sheet
            {
                Name = "Monthly Report",
                ColumnsWidth = new List<double> { 10, 12, 8, 8, 35 }
            };

            _excelWriter = new ExcelWriter(FilePath, sheet);
            for (var row = 1; row <= 100; row++)
            {
                for (var col = 1; col <= 10; col++)
                {
                    _excelWriter.Write($"row:{row}-col:{col}", col, row);
                }
            }

            _excelWriter.Save();


            //Initiate test collection
            var testCollection = new List<TestModel>
            {
                new TestModel(), new TestModel()
            };

            //Export list of objects to Excel file
            testCollection.ExportToExcel(FilePath);


            //Export list of objects to Excel file with predefined Sheet name
            testCollection.ExportToExcel(FilePath, sheetName: "Sheet2");
        }
    }
}