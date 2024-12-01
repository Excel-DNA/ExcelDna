using ExcelDna.Integration.ExtendedRegistration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;

namespace ExcelDna.Integration.ObjectHandles
{
    internal static class ObjectHandleRegistration
    {
        public static IEnumerable<ExcelFunction> ProcessObjectHandles(this IEnumerable<ExcelFunction> registrations)
        {
            registrations = registrations.ProcessParameterConversions(GetParameterConversionConfiguration());

            foreach (var reg in registrations)
            {
                if (HasExcelHandle(reg.Return.CustomAttributes))
                {
                    reg.FunctionLambda = LazyLambda.Create(reg.FunctionLambda);

                    EntryFunctionExecutionHandler entryFunctionExecutionHandler = new EntryFunctionExecutionHandler();

                    reg.FunctionLambda = FunctionExecutionRegistration.ApplyMethodHandler(reg.FunctionAttribute.Name, reg.FunctionLambda, entryFunctionExecutionHandler);

                    var returnConversion = CreateReturnConversion((object value) => GetReturnConversion(value, reg.FunctionAttribute.Name, entryFunctionExecutionHandler));

                    ParameterConversionRegistration.ApplyConversions(reg, null, ParameterConversionRegistration.GetReturnConversions(new List<ParameterConversionConfiguration.ReturnConversion> { returnConversion }, reg.FunctionLambda.ReturnType, reg.ReturnRegistration));
                }

                yield return reg;
            }
        }

        public static ParameterConversionConfiguration GetParameterConversionConfiguration()
        {
            return new ParameterConversionConfiguration().AddParameterConversion(GetParameterConversion());
        }

        public static bool IsMethodSupported(ExcelFunction reg)
        {
            if (HasExcelHandle(reg.Return.CustomAttributes))
                return true;

            return reg.Parameters.Any(paramReg => HasExcelHandle(paramReg.CustomAttributes));
        }

        public static int ProcessAssemblyAttributes(IEnumerable<object> attributes)
        {
            List<object> excelHandleAttribute = new List<object>();
            excelHandleAttribute.Add(new ExcelHandleAttribute());

            int result = 0;
            foreach (Type t in attributes.OfType<ExcelHandleExternalAttribute>().Select(i => i.Type))
            {
                ExcelTypeDescriptor.AddCustomAttributes(t, excelHandleAttribute);
                ++result;
            }

            return result;
        }

        static ParameterConversionConfiguration.ReturnConversion CreateReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            return new ParameterConversionConfiguration.ReturnConversion((unusedReturnType, unusedAttributes) => convert, null, false);
        }

        static object GetReturnConversion(object value, string callerFunctionName, EntryFunctionExecutionHandler entryFunctionExecutionHandler)
        {
            object result = ObjectHandler.GetHandle(callerFunctionName, entryFunctionExecutionHandler.GetArguments(Thread.CurrentThread.ManagedThreadId), value);

            return result;
        }

        static Func<Type, IExcelFunctionParameter, LambdaExpression> GetParameterConversion()
        {
            return (type, paramReg) => HandleStringConversion(type, paramReg);
        }

        static LambdaExpression HandleStringConversion(Type type, IExcelFunctionParameter paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!HasExcelHandle(paramReg.CustomAttributes))
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

            paramReg.CustomAttributes.RemoveAll(att => att is ExcelHandleAttribute);

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

        static bool HasExcelHandle(List<object> customAttributes)
        {
            return customAttributes.OfType<ExcelHandleAttribute>().Any();
        }

        private class EntryFunctionExecutionHandler : FunctionExecutionHandler
        {
            private ConcurrentDictionary<int, object> arguments = new ConcurrentDictionary<int, object>();

            public object GetArguments(int managedThreadId)
            {
                if (arguments.TryGetValue(managedThreadId, out object value))
                {
                    return value;
                }

                return null;
            }

            public override void OnEntry(FunctionExecutionArgs args)
            {
                this.arguments.AddOrUpdate(Thread.CurrentThread.ManagedThreadId, args.Arguments, (key, oldValue) => args.Arguments);
            }
        }
    }
}
