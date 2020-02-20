using System.Security;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        public ExcelWriter(string filePath, Sheet sheet = null)
            : base(filePath, sheet)
        {
        }

        public void Write(string value, int col, int row, DataType dataType = DataType.Text)
        {
            Sheet.PrepareRow(col, row);

            var data = GetCellData(value, col, row, dataType);
            Sheet.Write(data);

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
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
    }
}