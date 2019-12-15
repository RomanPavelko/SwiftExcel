using System.Collections.Generic;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        public ExcelWriter(string filePath, IList<Sheet> sheets)
            : base(filePath, sheets)
        {
        }
    }
}