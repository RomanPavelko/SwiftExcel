using SwiftExcel.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Specialized;
using System.IO;

namespace SwiftExcel.Sandbox
{
    internal class Program
    {
        private const string FilePath = "C:/temp/test.xlsx";
        private const string ConnectionString = "your_connection_string";
        private const string ContainerName = "your_container_name";
        private const string BlobName = "test.xlsx";

        private static ExcelWriter _excelWriter;

        private static void Main()
        {
            var stopwatch = Stopwatch.StartNew();

            //Fill excel document with test data 100 rows x 10 columns
            using (_excelWriter = new ExcelWriter(FilePath))
            {
                for (var row = 1; row <= 100; row++)
                {
                    for (var col = 1; col <= 100; col++)
                    {
                        _excelWriter.Write($"row:{row}-col:{col}", col, row);
                    }
                }
            }


            //Use Azure Blob Storage to directly upload file over a stream
            //var blobServiceClient = new BlobServiceClient(ConnectionString);
            //var containerClient = blobServiceClient.GetBlobContainerClient(ContainerName);
            //var blobClient = containerClient.GetBlockBlobClient(BlobName);
            //using (var stream = blobClient.OpenWrite(true))
            //{
            //    using (_excelWriter = new ExcelWriter(stream))
            //    {
            //        for (var row = 1; row <= 100; row++)
            //        {
            //            for (var col = 1; col <= 100; col++)
            //            {
            //                _excelWriter.Write($"row:{row}-col:{col}", col, row);
            //            }
            //        }
            //    }
            //}


            //Upload generated file to Azure Blob Storage
            //using (_excelWriter = new ExcelWriter(FilePath))
            //{
            //    for (var row = 1; row <= 100; row++)
            //    {
            //        for (var col = 1; col <= 100; col++)
            //        {
            //            _excelWriter.Write($"row:{row}-col:{col}", col, row);
            //        }
            //    }
            //}
            //var blobServiceClient = new BlobServiceClient(ConnectionString);
            //var containerClient = blobServiceClient.GetBlobContainerClient(ContainerName);
            //var blobClient = containerClient.GetBlockBlobClient(BlobName);
            //using (var fileStream = new FileStream(FilePath, FileMode.Open))
            //{
            //    blobClient.Upload(fileStream);
            //}


            //Invalid XML characters
            //Configuration.UseEnchancedXmlEscaping = true;
            //using (_excelWriter = new ExcelWriter(FilePath))
            //{
            //    _excelWriter.Write("<", 1, 1);
            //    _excelWriter.Write(">", 2, 1);
            //    _excelWriter.Write("&", 3, 1);
            //    _excelWriter.Write("'", 4, 1);
            //    _excelWriter.Write("\"", 5, 1);
            //}


            //Set custom sheet name, define columns width, right to left and wrap text
            //Use manual Save() instead of using block 
            //var sheet = new Sheet
            //{
            //    Name = "Monthly Report",
            //    RightToLeft = true,
            //    WrapText = true,
            //    ColumnsWidth = new List<double> { 10, 12, 8, 8, 35 }
            //};

            //_excelWriter = new ExcelWriter(FilePath, sheet);
            //for (var row = 1; row <= 100; row++)
            //{
            //    for (var col = 1; col <= 10; col++)
            //    {
            //        _excelWriter.Write($"row:{row}-col:{col}", col, row);
            //    }
            //}

            //_excelWriter.Save();


            ////Formula examples
            //using (_excelWriter = new ExcelWriter(FilePath))
            //{
            //    const int col = 1;
            //    var row = 1;
            //    for (; row <= 20; row++)
            //    {
            //        _excelWriter.Write(row.ToString(), col, row, DataType.Number);
            //    }

            //    _excelWriter.WriteFormula(FormulaType.Average, col, ++row, col, 1, 20);
            //    _excelWriter.WriteFormula(FormulaType.Count, col, ++row, col, 1, 20);
            //    _excelWriter.WriteFormula(FormulaType.Max, col, ++row, col, 1, 20);
            //    _excelWriter.WriteFormula(FormulaType.Sum, col, ++row, col, 1, 20);
            //}


            ////Initiate test collection
            //var testCollection = new List<TestModel>
            //{
            //    new TestModel(), new TestModel()
            //};

            ////Export list of objects to Excel file
            //testCollection.ExportToExcel(FilePath);


            ////Export list of objects to Excel file with predefined Sheet name
            //testCollection.ExportToExcel(FilePath, sheetName: "Sheet2");

            stopwatch.Stop();
            Console.WriteLine($"Completed in {stopwatch.ElapsedMilliseconds} ms.");
        }
    }
}