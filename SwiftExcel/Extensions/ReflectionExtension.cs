using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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
        
        public static DisplayAttribute GetDisplayAttributeOrDefault(this MemberInfo memberInfo)
        {
            var displayAttributes = memberInfo.GetCustomAttributes(typeof(DisplayAttribute), true);

            return displayAttributes.Any() ? displayAttributes.Cast<DisplayAttribute>().Single() : null;
        }

        public static string GetDisplayName(this MemberInfo memberInfo)
        {
            var displayName = memberInfo.GetDisplayAttributeOrDefault()?.GetName();
            return !string.IsNullOrEmpty(displayName) ? displayName : memberInfo.Name;
        }

        public static int GetDisplayOrder(this MemberInfo memberInfo)
        {
            return memberInfo.GetDisplayAttributeOrDefault()?.GetOrder() ?? int.MaxValue;
        }

        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);
        }
    }
}
