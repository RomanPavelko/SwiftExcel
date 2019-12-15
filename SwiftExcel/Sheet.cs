using System.Collections.Generic;

namespace SwiftExcel
{
    public class Sheet
    {
        internal string Path { get; set; }

        public string Name { get; set; }
        public IList<double> ColumnsWidth { get; set; }
    }
}