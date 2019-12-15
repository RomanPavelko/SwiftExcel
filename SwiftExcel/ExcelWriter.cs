using System.Collections.Generic;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        public ExcelWriter(string filePath, IList<string> sheetNames)
            : base(filePath, sheetNames)
        {
        }
    }
}