using System;

namespace SwiftExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelExportAttribute: Attribute
    {
        private int? order;

        private double? width;

        public string Name { get; set; }

        public int Order
        {
            get => order.GetValueOrDefault();
            set => order = value;
        }

        public double Width
        {
            get => width.GetValueOrDefault();
            set => width = value;
        }

        public string GetName()
        {
            return !string.IsNullOrEmpty(Name) ? Name : null;
        }

        public int? GetOrder()
        {
            return order;
        }

        public double? GetWidth()
        {
            return width;
        }
    }
}
