﻿using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

namespace SwiftExcel
{
    public class ExcelWriter : ExcelWriterCore
    {
        /// <summary>
        /// Create an instance of ExcelWriter
        /// </summary>
        /// <param name="filePath">Full path to the result file</param>
        /// <param name="sheet">Optional: Custom sheet configuration</param>
        public ExcelWriter(string filePath, Sheet sheet = null)
            : base(filePath, sheet)
        {
        }

        /// <summary>
        /// Create an instance of ExcelWriter
        /// </summary>
        /// <param name="stream">Provide your own stream to the result file. It can be either your own FileStream or a stream representing a cloud storage instance like Azure Blob</param>
        /// <param name="sheet">Optional: Custom sheet configuration</param>
        public ExcelWriter(Stream stream, Sheet sheet = null)
            : base(stream, sheet)
        {
        }

        public void Write(string value, int col, int row, DataType dataType = DataType.Text)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write("<c");
            if (dataType == DataType.Text)
            {
                Sheet.Write(" t=\"str\"");
            }
            Sheet.Write("><v>");
            Sheet.Write(EscapeInvalidChars(value));
            Sheet.Write("</v></c>");

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
        }

        public void WriteFormula(FormulaType type, int col, int row, int sourceCol, int sourceRowStart, int sourceRowEnd)
        {
            Sheet.PrepareRow(col, row);

            Sheet.Write("<c><f>");
            Sheet.Write($"{type.ToString().ToUpper()}");
            Sheet.Write("(");
            Sheet.Write($"{GetFullCellName(sourceCol, sourceRowStart)}");
            Sheet.Write(":");
            Sheet.Write($"{GetFullCellName(sourceCol, sourceRowEnd)}");
            Sheet.Write(")</f><v></v></c>");

            Sheet.CurrentCol = col;
            Sheet.CurrentRow = row;
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
            if (!Configuration.UseEnchancedXmlEscaping)
            {
                return value;
            }

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