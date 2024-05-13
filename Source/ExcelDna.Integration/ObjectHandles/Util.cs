using ExcelDna.Integration.ExtendedRegistration;
using System;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class Util
    {
        static ObjectHandler _objectHandler = new ObjectHandler();

        public static object ReturnConversionNew(object value, string callerFunctionName, object callerParameters)
        {
            bool newHandle;
            object result = _objectHandler.GetHandle(callerFunctionName, callerParameters, value, out newHandle);
            if (!newHandle)
                (value as IDisposable)?.Dispose();

            return result;
        }

        public static Func<Type, ExcelParameter, LambdaExpression> GetParameterConversion()
        {
            return (type, paramReg) => HandleStringConversionNew(type, paramReg);
        }

        static LambdaExpression HandleStringConversionNew(Type type, ExcelParameter paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(ExcelObjectHandle<>))
                return null;

            //return null;
            var input = Expression.Parameter(typeof(object), "input");
            var objectType = typeof(object);
            Expression<Func<Type, object, object>> enumParse = (t, s) => GetObjectNew((string)s);
            var result =
                Expression.Lambda(
                    Expression.Convert(
                        Expression.Invoke(enumParse, Expression.Constant(type), input),
                        type),
                    input);
            return result;
        }

        static object GetObjectNew(string handle)
        {
            object value;
            // TODO: We might be able to strongly type the GetObject...
            if (_objectHandler.TryGetObject(handle, out value))
            {
                return value;
            }
            // No object for the handle ...
            return "!!! INVALID HANDLE";
        }
    }
}
