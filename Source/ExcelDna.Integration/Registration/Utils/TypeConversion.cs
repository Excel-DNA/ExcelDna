using ExcelDna.Integration;
using System;
using System.Linq.Expressions;

namespace ExcelDna.Registration
{
    public static class TypeConversion
    {
        public static LambdaExpression GetConversion(Type inputType, Type targetType)
        {
            var input = Expression.Parameter(typeof(object), "input");
            return Expression.Lambda(GetConversion(input, targetType), input);
        }

        public static Expression GetConversion(Expression input, Type type)
        {
            if (type == typeof(Object))
                return input;
            if (type == typeof(Double))
                return Expression.Call(((Func<Object, Double>)ConvertToDouble).Method, input);
            if (type == typeof(String))
                return Expression.Call(((Func<Object, String>)ConvertToString).Method, input);
            if (type == typeof(DateTime))
                return Expression.Call(((Func<Object, DateTime>)ConvertToDateTime).Method, input);
            if (type == typeof(Boolean))
                return Expression.Call(((Func<Object, Boolean>)ConvertToBoolean).Method, input);
            if (type == typeof(Int64))
                return Expression.Call(((Func<Object, Int64>)ConvertToInt64).Method, input);
            if (type == typeof(Int32))
                return Expression.Call(((Func<Object, Int32>)ConvertToInt32).Method, input);
            if (type == typeof(Int16))
                return Expression.Call(((Func<Object, Int16>)ConvertToInt16).Method, input);
            if (type == typeof(UInt16))
                return Expression.Call(((Func<Object, UInt16>)ConvertToUInt16).Method, input);
            if (type == typeof(Decimal))
                return Expression.Call(((Func<Object, Decimal>)ConvertToDecimal).Method, input);

            // Fallback - not likely to be useful
            return Expression.Convert(input, type);
        }

        public static double ConvertToDouble(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (double)result;
            }

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Int32.");
        }

        public static string ConvertToString(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeString);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (string)result;
            }

            // Not sure how this can happen...
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to String.");
        }

        public static string[,] ConvertToString2D(object[,] values)
        {
            string[,] result = new string[values.GetLength(0), values.GetLength(1)];
            for (int i = 0; i < values.GetLength(0); i++)
            {
                for (int j = 0; j < values.GetLength(1); j++)
                {
                    result[i, j] = ConvertToString(values[i, j]);
                }
            }

            return result;
        }

        public static DateTime ConvertToDateTime(object value)
        {
            try
            {
                return DateTime.FromOADate(ConvertToDouble(value));
            }
            catch
            {
                // Might exceed range of DateTime
                throw new InvalidCastException("Value " + value.ToString() + " could not be converted to DateTime.");
            }
        }

        public static bool ConvertToBoolean(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeBoolean);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return (bool)result;

            // failed - as a fallback, try to convert to a double
            retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return ((double)result != 0.0);

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Boolean.");
        }

        public static int ConvertToInt32(object value)
        {
            return checked((int)ConvertToInt64(value));
        }

        public static short ConvertToInt16(object value)
        {
            return checked((short)ConvertToInt64(value));
        }

        public static ushort ConvertToUInt16(object value)
        {
            return checked((ushort)ConvertToInt64(value));
        }

        public static decimal ConvertToDecimal(object value)
        {
            return checked((decimal)ConvertToDouble(value));
        }

        public static long ConvertToInt64(object value)
        {
            return checked((long)Math.Round(ConvertToDouble(value), MidpointRounding.ToEven));
        }

        public static object GetDefault(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }
    }

    internal enum XlType : int
    {
        XlTypeNumber = 0x0001,
        XlTypeString = 0x0002,
        XlTypeBoolean = 0x0004,
        XlTypeReference = 0x0008,
        XlTypeError = 0x0010,
        XlTypeFlow = 0x0020, // Unused
        XlTypeArray = 0x0040,
        XlTypeMissing = 0x0080,
        XlTypeEmpty = 0x0100,
        XlTypeInt = 0x0800,     // int16 in XlOper, int32 in XlOper12, never passed into UDF
    }
}
