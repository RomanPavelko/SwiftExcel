using System;

namespace SwiftExcel.Exceptions
{
    public class SwiftExcelException : Exception
    {
        public SwiftExcelException(SwiftExcelExceptionType type, object data = null, object additionalData = null)
            : base(GetMessage(type, data, additionalData))
        {
        }

        internal static string GetMessage(SwiftExcelExceptionType type, object data, object additionalData)
        {
            switch (type)
            {
                case SwiftExcelExceptionType.SheetNumberLessThanOne:
                    return "Sheet number must be 1 or greater.";
                case SwiftExcelExceptionType.SheetNumberOutOfRange:
                    return $"Sheet with number {data} was not defined.";
                case SwiftExcelExceptionType.ColNumberLessThanOne:
                    return "Column number must be 1 or greater.";
                case SwiftExcelExceptionType.RowNumberLessThanOne:
                    return "Row number must be 1 or greater.";
                case SwiftExcelExceptionType.RowNumberAlreadyProcessed:
                    return $"Row {data} has already been processed.";
                case SwiftExcelExceptionType.ColNumberAlreadyProcessed:
                    return $"Column {data} has already been processed in Row {additionalData}.";
                default:
                    return "Unhandled exception";
            }
        }
    }
}