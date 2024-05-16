using ExcelDna.Integration.ExtendedRegistration;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ObjectHandles
{
    internal static class ObjectHandleRegistration
    {
        public static IEnumerable<ExcelFunction> ProcessObjectHandles(this IEnumerable<ExcelFunction> registrations)
        {
            var paramConversionConfig = new ParameterConversionConfiguration().AddParameterConversion(GetParameterConversion());
            registrations = registrations.ProcessParameterConversions(paramConversionConfig);

            foreach (var reg in registrations)
            {
                if (IsExcelObjectHandle(reg.FunctionLambda.ReturnType))
                {
                    EntryFunctionExecutionHandler entryFunctionExecutionHandler = new EntryFunctionExecutionHandler();

                    reg.FunctionLambda = FunctionExecutionRegistration.ApplyMethodHandler(reg.FunctionAttribute.Name, reg.FunctionLambda, entryFunctionExecutionHandler);

                    var returnConversion = CreateReturnConversion((object value) => GetReturnConversion(value, reg.FunctionAttribute.Name, entryFunctionExecutionHandler.Args.Arguments));

                    ParameterConversionRegistration.ApplyConversions(reg, null, ParameterConversionRegistration.GetReturnConversions(new List<ParameterConversionConfiguration.ReturnConversion> { returnConversion }, reg.FunctionLambda.ReturnType, reg.ReturnRegistration));
                }

                yield return reg;
            }
        }

        static ParameterConversionConfiguration.ReturnConversion CreateReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            return new ParameterConversionConfiguration.ReturnConversion((unusedReturnType, unusedAttributes) => convert, null, false);
        }

        static object GetReturnConversion(object value, string callerFunctionName, object callerParameters)
        {
            bool newHandle;
            object result = ObjectHandler.GetHandle(callerFunctionName, callerParameters, value, out newHandle);
            if (!newHandle)
                (value as IDisposable)?.Dispose();

            return result;
        }

        static Func<Type, ExcelParameter, LambdaExpression> GetParameterConversion()
        {
            return (type, paramReg) => HandleStringConversion(type, paramReg);
        }

        static LambdaExpression HandleStringConversion(Type type, ExcelParameter paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!IsExcelObjectHandle(type))
                return null;

            var input = Expression.Parameter(typeof(object), "input");
            var objectType = typeof(object);
            Expression<Func<Type, object, object>> parse = (t, s) => GetObject((string)s);
            var result =
                Expression.Lambda(
                    Expression.Convert(
                        Expression.Invoke(parse, Expression.Constant(type), input),
                        type),
                    input);
            return result;
        }

        static object GetObject(string handle)
        {
            object value;
            if (ObjectHandler.TryGetObject(handle, out value))
            {
                return value;
            }

            // No object for the handle ...
            return "!!! INVALID HANDLE";
        }

        static bool IsExcelObjectHandle(Type t)
        {
            return t.IsGenericType && t.GetGenericTypeDefinition() == typeof(ExcelObjectHandle<>);
        }

        private class EntryFunctionExecutionHandler : FunctionExecutionHandler
        {
            public FunctionExecutionArgs Args { get; private set; }

            public override void OnEntry(FunctionExecutionArgs args)
            {
                this.Args = args;
            }
        }
    }
}
