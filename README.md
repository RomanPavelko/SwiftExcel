# SwiftExcel
Lightweight, extremely fast and memory efficient Excel output library for .NET and .NET Core applications. Build your Excel reports in fraction of seconds with no memory footprint thanks to skipping XML serialization and streaming data directly to the file.
# Installation
SwiftExcel is available as a NuGet package. You can install it using the NuGet Package Console window:
```
PM> Install-Package SwiftExcel
```
# Usage
### Fill excel document with test data 100 rows x 10 columns
```
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
```
var sheet = new Sheet
{
    Name = "Monthly Report", 
    ColumnsWidth = new List<double> { 10, 12, 8, 8, 35 }
};

var ew = new ExcelWriter("C:\\temp\\test.xlsx", sheet)
```
