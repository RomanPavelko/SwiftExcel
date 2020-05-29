using System;
using System.ComponentModel.DataAnnotations;

namespace SwiftExcel.Sandbox
{
    public class TestModel
    {
        [Display(Order = 1)]
        public int FirstProperty { get; set; } = int.MaxValue;

        [Display(Name = "Custom SecondProperty Name", Order = 2)]
        public string SecondProperty { get; set; } = nameof(SecondProperty);

        [Display(Name = "Custom ThirdProperty Name")]
        public char ThirdProperty { get; set; } = char.MaxValue;

        public DateTime FourthProperty { get; set; } = DateTime.MaxValue;
    }
}
