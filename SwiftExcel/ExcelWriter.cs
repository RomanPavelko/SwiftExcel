using System.Collections.Generic;
using System.Security;
using SwiftExcel.Exceptions;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        public ExcelWriter(string filePath, IList<Sheet> sheets)
            : base(filePath, sheets)
        {
        }

        public void Write(string value, int col, int row, int sheetNumber, DataType dataType = DataType.Text)
        {
            var sheet = GetSheet(sheetNumber);
            sheet.PrepareRow(col, row);

            var data = GetCellData(value, col, row, dataType);
            sheet.Write(data);

            sheet.CurrentRow = row;
        }

        private static string GetCellData(string value, int col, int row, DataType dataType)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var t = dataType == DataType.Text ? " t=\"str\"" : string.Empty;
            return $"<c r=\"{GetFullCellName(col, row)}\"{t}><v>{SecurityElement.Escape(value.Trim())}</v></c>";
        }
        
        private static string GetFullCellName(int col, int row)
        {
            return $"{GetCellName(col)}{row}";
        }

        private static string GetCellName(int col)
        {
            var dividend = col;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = (char)(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private Sheet GetSheet(int sheetNumber)
        {
            if (sheetNumber <= 0)
            {
                throw new SwiftExcelException(SwiftExcelExceptionType.SheetNumberLessThanOne);
            }
            if (sheetNumber > Sheets.Count)
            {
                throw new SwiftExcelException(SwiftExcelExceptionType.SheetNumberOutOfRange, sheetNumber);
            }

            return Sheets[sheetNumber - 1];
        }
    }
}