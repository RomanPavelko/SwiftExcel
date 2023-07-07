# SwiftExcel
[![Official Site](https://img.shields.io/badge/site-swiftexcel-blue.svg)](https://swiftexcel.pro/) [![Latest version](https://img.shields.io/nuget/v/SwiftExcel.svg)](https://www.nuget.org/packages?q=SwiftExcel) [![License MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://en.wikipedia.org/wiki/MIT_License)
# Overview
Lightweight, extremely fast and memory efficient Excel output library for .NET and .NET Core applications. Build your Excel reports in fraction of seconds with no memory footprint thanks to skipping XML serialization and streaming data directly to the file.
# Installation
SwiftExcel is available as a NuGet package. You can install it using the NuGet Package Console window:
```
PM> Install-Package SwiftExcel
```
# Usage
### Fill excel document with test data 100 rows x 10 columns
```csharp
using (var ew = new ExcelWriter("C:\\temp\\test.xlsx"))
{
    for (var row = 1; row <= 100; row++)
    {
        for (var col = 1; col <= 10; col++)
        {
            ew.Write($"row:{row}-col:{col}", col, row);
        }
    }
}
```
### Set custom sheet name and define columns width
```csharp
var sheet = new Sheet
{
    Name = "Monthly Report", 
    ColumnsWidth = new List<double> { 10, 12, 8, 8, 35 }
};

var ew = new ExcelWriter("C:\\temp\\test.xlsx", sheet);
```
### Using formulas. Output 20 values and calculate average, count, max and sum
```csharp
using (var ew = new ExcelWriter("C:\\temp\\test.xlsx"))
{
    for (var row = 1; row <= 20; row++)
    {
        ew.Write(row.ToString(), 1, row, DataType.Number);
    }

    ew.WriteFormula(FormulaType.Average, 1, 22, 1, 1, 20);
    ew.WriteFormula(FormulaType.Count, 1, 23, 1, 1, 20);
    ew.WriteFormula(FormulaType.Max, 1, 24, 1, 1, 20);
    ew.WriteFormula(FormulaType.Sum, 1, 25, 1, 1, 20);
}
```
# Performance
SwiftExcel has incredible performance due to ignoring XML serialization and streaming data directly to the file. 
Below is a performance test creating a document with 100 000 rows and 100 columns compared to other popular Excel output libraries on Nuget.

|   | Execution Time | Memory Usage |
| :--- | :---: | :---: |
| SwiftExcel  | 6.1 sec  |  19 mb  |
| FastExcel  | 31.1 sec  |  3200 mb  |
| EPPlus  | 44.2 sec  |  2900 mb  |
| Syncfusion.XlsIO  | 73.3 sec  |  2700 mb  |
| IronXL.Excel  | 306.8 sec  |  7700 mb  |
| Microsoft Interop Excel  | >3 hours  |  27 mb  |

# SwiftExcel.Pro
Check out the [.Pro version](https://swiftexcel.pro/) for advanced features like working with multiple sheets or applying custom styles.
