using ExcelService.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelService.Helper
{
    internal class ExcelHelper
    {
        internal static T ConvertToType<T>(object value)
        {
            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
            {
                return GetDefaultValue<T>(typeof(T));
            }

            var targetType = typeof(T);

            var underlyingType = Nullable.GetUnderlyingType(targetType);
            if (underlyingType != null)
            {
                return (T)ConvertToTypeDynamic(value, underlyingType);
            }

            return (T)ConvertToTypeDynamic(value, targetType);
        }

        private static object ConvertToTypeDynamic(object value, System.Type targetType)
        {
            if (value != null && targetType.IsAssignableFrom(value.GetType()))
            {
                return value;
            }

            string strValue = value?.ToString() ?? string.Empty;

            switch (System.Type.GetTypeCode(targetType))
            {
                case System.TypeCode.String:
                    return strValue;
                case System.TypeCode.Boolean:
                    return bool.TryParse(strValue, out var b) ? b : default(bool);
                case System.TypeCode.Byte:
                    return byte.TryParse(strValue, out var bt) ? bt : default(byte);
                case System.TypeCode.Char:
                    return char.TryParse(strValue, out var c) ? c : default(char);
                case System.TypeCode.Int16:
                    return short.TryParse(strValue, out var s) ? s : default(short);
                case System.TypeCode.UInt16:
                    return ushort.TryParse(strValue, out var us) ? us : default(ushort);
                case System.TypeCode.Int32:
                    return int.TryParse(strValue, out var i) ? i : default(int);
                case System.TypeCode.UInt32:
                    return uint.TryParse(strValue, out var ui) ? ui : default(uint);
                case System.TypeCode.Int64:
                    return long.TryParse(strValue, out var l) ? l : default(long);
                case System.TypeCode.UInt64:
                    return ulong.TryParse(strValue, out var ul) ? ul : default(ulong);
                case System.TypeCode.Single:
                    return float.TryParse(strValue, out var f) ? f : default(float);
                case System.TypeCode.Double:
                    return double.TryParse(strValue, out var d) ? d : default(double);
                case System.TypeCode.Decimal:
                    return decimal.TryParse(strValue, out var dec) ? dec : default(decimal);
                case System.TypeCode.DateTime:
                    return DateTime.TryParse(strValue, out var dt) ? dt : default(DateTime);
                case System.TypeCode.Object:
                    if (targetType == typeof(System.Guid))
                        return System.Guid.TryParse(strValue, out var g) ? g : System.Guid.Empty;

                    if (targetType == typeof(System.DateTimeOffset))
                        return System.DateTimeOffset.TryParse(strValue, out var dto) ? dto : default(System.DateTimeOffset);

                    if (targetType == typeof(System.TimeSpan))
                        return System.TimeSpan.TryParse(strValue, out var ts) ? ts : default(System.TimeSpan);

                    if (targetType.IsEnum)
                    {
                        try
                        {
                            return System.Enum.Parse(targetType, strValue, ignoreCase: true);
                        }
                        catch
                        {
                            return System.Activator.CreateInstance(targetType);
                        }
                    }
                    return System.Convert.ChangeType(value, targetType);
                default:
                    return System.Convert.ChangeType(value, targetType);
            }
        }

        private static T GetDefaultValue<T>(System.Type type)
        {
            if (!type.IsValueType)
            {
                return default(T);
            }
            return System.Activator.CreateInstance<T>();
        }

    }
}
