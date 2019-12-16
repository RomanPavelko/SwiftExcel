using System.Collections.Generic;
using System.IO;
using SwiftExcel.Exceptions;

namespace SwiftExcel
{
    public class Sheet
    {
        public string Name { get; set; }
        public IList<double> ColumnsWidth { get; set; }

        internal TextWriter TextWriter { get; set; }
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
        }
    }
}