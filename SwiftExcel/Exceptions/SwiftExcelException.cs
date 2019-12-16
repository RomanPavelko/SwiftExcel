using System;

namespace SwiftExcel.Exceptions
{
    public class SwiftExcelException : Exception
    {
        public SwiftExcelException(SwiftExcelExceptionType type, object data = null)
            : base(GetMessage(type, data))
        {
        }

        internal static string GetMessage(SwiftExcelExceptionType type, object data)
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
                default:
                    return "Unhandled exception";
            }
        }
    }
}