using SwiftExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SwiftExcel.Extensions
{
    public static class ReflectionExtension
    {
        private static readonly IList<Type> NumericTypes = new List<Type>
        {
            typeof(int),  typeof(double),  typeof(decimal),
            typeof(long), typeof(short),   typeof(sbyte),
            typeof(byte), typeof(ulong),   typeof(ushort),
            typeof(uint), typeof(float)
        };

        public static string GetPropertyValue<TData>(this PropertyInfo propertyInfo, TData entity)
        {
            return propertyInfo.GetValue(entity)?.ToString();
        }

        public static ExcelExportAttribute GetExcelExportAttributeOrDefault(this MemberInfo memberInfo)
        {
            var exportAttributes = memberInfo.GetCustomAttributes(typeof(ExcelExportAttribute), true);

            return exportAttributes.Any() ? exportAttributes.Cast<ExcelExportAttribute>().Single() : null;
        }

        public static string GetExportNameOrDefault(this MemberInfo memberInfo)
        {
            return memberInfo.GetExcelExportAttributeOrDefault()?.GetName();
        }

        public static int? GetExportOrderOrDefault(this MemberInfo memberInfo)
        {
            return memberInfo.GetExcelExportAttributeOrDefault()?.GetOrder();
        }

        public static double? GetExportWidthOrDefault(this MemberInfo memberInfo)
        {
            return memberInfo.GetExcelExportAttributeOrDefault()?.GetWidth();
        }

        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);
        }
    }
}
