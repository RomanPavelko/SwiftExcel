using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

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
            var data = GetCellData(value, col, row, dataType);
            Write(col, row, data);
        }

        public void WriteFormula(FormulaType type, int col, int row, int sourceCol, int sourceRowStart, int sourceRowEnd)
        {
            var data = GetCellDataFormula(type, col, row, sourceCol, sourceRowStart, sourceRowEnd);
            Write(col, row, data);
        }

        private void Write(int col, int row, string data)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write(data);

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }

        private static string GetCellData(string value, int col, int row, DataType dataType)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var t = dataType == DataType.Text ? " t=\"str\"" : string.Empty;
            return $"<c r=\"{GetFullCellName(col, row)}\"{t}><v>{EscapeInvalidChars(value.Trim())}</v></c>";
        }

        private static string GetCellDataFormula(FormulaType type, int col, int row, int sourceCol, int sourceRowStart, int sourceRowEnd)
        {
            return $"<c r=\"{GetFullCellName(col, row)}\"><f>{type.ToString().ToUpper()}({GetFullCellName(sourceCol, sourceRowStart)}:{GetFullCellName(sourceCol, sourceRowEnd)})</f><v></v></c>";
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

        private static string EscapeInvalidChars(string value)
        {
            value = SecurityElement.Escape(value);

            if (string.IsNullOrEmpty(value) || value.All(XmlConvert.IsXmlChar))
            {
                return value;
            }

            var result = new StringBuilder();
            foreach (var character in value)
            {
                if (XmlConvert.IsXmlChar(character))
                {
                    result.Append(character);
                }
                else
                {
                    result.Append($"_x{(int)character:x4}_");
                }
            }

            return result.ToString();
        }
    }
}