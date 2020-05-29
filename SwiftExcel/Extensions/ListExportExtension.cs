using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SwiftExcel.Extensions
{
    public static class ListExportExtension
    {
        private const int HeaderRowNumber = 1;

        private const int SkipZeroIndexStep = 1;

        public static void ExportToExcel<TData>(this IList<TData> entities, string filePath, string sheetName = null, double? columnsWidth = null)
        {
            var properties = typeof(TData).GetProperties().Order();

            using (var excelWriter = new ExcelWriter(filePath, properties.CreateSheet(sheetName, columnsWidth)))
            {
                try
                {
                    excelWriter
                        .CreateHeader(properties)
                        .PopulateBody(entities, properties)
                        .Save();
                }
                finally
                {
                    excelWriter.Dispose();
                }
            }
        }

        private static IList<PropertyInfo> Order(this IList<PropertyInfo> properties)
        {
            return properties.OrderBy(property => property.GetDisplayOrder()).ToList();
        }

        private static Sheet CreateSheet(this IList<PropertyInfo> properties, string sheetName, double? columnsWidth)
        {
            return new Sheet
            {
                Name = sheetName ?? Sheet.DefaultName,
                ColumnsWidth = columnsWidth.HasValue ? properties.Select(property => columnsWidth.Value).ToList() : default
            };
        }

        private static ExcelWriter CreateHeader(this ExcelWriter excelWriter, IList<PropertyInfo> properties)
        {
            for (var i = 0; i < properties.Count; i++)
            {
                var columnNumber = i + SkipZeroIndexStep;
                excelWriter.Write(properties[i].GetDisplayName(), columnNumber, HeaderRowNumber);
            }

            return excelWriter;
        }

        private static ExcelWriter PopulateBody<TData>(this ExcelWriter excelWriter, IList<TData> entities, IList<PropertyInfo> properties)
        {
            for (var i = 0; i < entities.Count; i++)
            {
                for (var j = 0; j < properties.Count; j++)
                {
                    var columnNumber = j + SkipZeroIndexStep;
                    var rowNumber = i + SkipZeroIndexStep + HeaderRowNumber;

                    excelWriter.Write(
                        properties[j].GetPropertyValue(entities[i]),
                        columnNumber,
                        rowNumber,
                        properties[j].PropertyType.GetExcelType());
                }
            }

            return excelWriter;
        }

        private static DataType GetExcelType(this Type type)
        {
            return type.IsNumeric() ? DataType.Number : DataType.Text;
        }
    }
}
