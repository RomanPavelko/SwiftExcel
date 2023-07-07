using SwiftExcel.Exceptions;
using System.Collections.Generic;
using System.IO;

namespace SwiftExcel
{
    public class Sheet
    {
        public const string DefaultName = "Sheet1";

        public string Name { get; set; } = DefaultName;
        public bool RightToLeft { get; set; }
        public bool WrapText { get; set; }
        public IList<double> ColumnsWidth { get; set; }

        internal Stream Stream { get; set; }
        internal StreamWriter StreamWriter { get; set; }
        internal int CurrentCol { get; set; }
        internal int CurrentRow { get; set; }

        internal void Write(string value)
        {
            StreamWriter.Write(value);
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

            //fill empty cells
            var colDifference = col - CurrentCol;
            if (colDifference > 1)
            {
                for (var i = 1; i < colDifference; i++)
                {
                    Write("<c t=\"str\"><v></v></c>");
                }
            }
        }

        internal string GetFormattedName()
        {
            return string.IsNullOrEmpty(Name)
                ? DefaultName
                : Name.Length > 31
                    ? Name.Substring(0, 31).Trim()
                    : Name;
        }
    }
}