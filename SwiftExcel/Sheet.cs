using SwiftExcel.Exceptions;
using System.Collections.Generic;
using System.IO;

namespace SwiftExcel
{
    public class Sheet
    {
        public const string DefaultName = "Sheet1";

        public string Name { get; set; } = DefaultName;
        public IList<double> ColumnsWidth { get; set; }

        internal TextWriter TextWriter { get; set; }
        internal int CurrentCol { get; set; }
        internal int CurrentRow { get; set; }

        internal void Write(string value)
        {
            TextWriter.Write(value);
        }

        internal void PrepareRow(int col, int row)
        {
            if (col <= 0)
            {
                throw new SwiftExcelException(SwiftExcelExceptionType.ColNumberLessThanOne, row);
            }
            if (row <= 0)
            {
                throw new SwiftExcelException(SwiftExcelExceptionType.RowNumberLessThanOne, row);
            }

            if (row > CurrentRow)
            {
                //close previous row
                if (CurrentRow != 0)
                {
                    Write("</row>");
                }

                //if skipping rows - add empty entries
                if (row - CurrentRow > 1)
                {
                    for (var i = 1; i < row - CurrentRow; i++)
                    {
                        Write("<row/>");
                    }
                }

                Write("<row>");
            }
            else if (row < CurrentRow)
            {
                throw new SwiftExcelException(SwiftExcelExceptionType.RowNumberAlreadyProcessed, row);
            }
            else
            {
                if (col <= CurrentCol)
                {
                    throw new SwiftExcelException(SwiftExcelExceptionType.ColNumberAlreadyProcessed, col, row);
                }
            }
        }
    }
}