using SwiftExcel.Attributes;
using System;

namespace SwiftExcel.Sandbox
{
    public class TestModel
    {
        [ExcelExport(Order = 1)]
        public int FirstProperty { get; set; } = int.MaxValue;

        [ExcelExport(Name = "Custom SecondProperty Name", Order = 2)]
        public string SecondProperty { get; set; } = nameof(SecondProperty);

        [ExcelExport(Name = "Custom ThirdProperty Name", Order = 3, Width = 40.00)]
        public char ThirdProperty { get; set; } = char.MaxValue;

        [ExcelExport(Width = 60.00)]
        public DateTime FourthProperty { get; set; } = DateTime.MaxValue;

        public byte FifthProperty { get; set; } = byte.MaxValue;
    }
}
