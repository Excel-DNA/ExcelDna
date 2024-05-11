using ExcelDna.Integration.ExtendedRegistration;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class MyFunctionExecutionHandler : FunctionExecutionHandler
    {
        public FunctionExecutionArgs args;

        public override void OnEntry(FunctionExecutionArgs args)
        {
            this.args = args;
        }
    }

    internal static class ObjectHandleRegistration
    {
        public static IEnumerable<ExcelFunction> ProcessObjectHandles(this IEnumerable<ExcelFunction> registrations)
        {
            var paramConversionConfig = new ParameterConversionConfiguration().AddParameterConversion(Util.GetParameterConversion());
            registrations = registrations.ProcessParameterConversions(paramConversionConfig);

            foreach (var reg in registrations)
            {
                if (reg.FunctionLambda.ReturnType.IsGenericType && reg.FunctionLambda.ReturnType.GetGenericTypeDefinition() == typeof(ExcelObjectHandle<>))
                {
                    MyFunctionExecutionHandler myFunctionExecutionHandler = new MyFunctionExecutionHandler();

                    reg.FunctionLambda = FunctionExecutionRegistration.ApplyMethodHandler(reg.FunctionAttribute.Name, reg.FunctionLambda, myFunctionExecutionHandler);

                    var rc = CreateReturnConversion((object value) => ObjectHandles.Util.ReturnConversionNew(value, reg.FunctionAttribute.Name, myFunctionExecutionHandler.args.Arguments));
                    var r = new List<ParameterConversionConfiguration.ReturnConversion> { rc };

                    List<LambdaExpression> rcs = ParameterConversionRegistration.GetReturnConversions(r, reg.FunctionLambda.ReturnType, reg.ReturnRegistration);

                    ParameterConversionRegistration.ApplyConversions(reg, null, rcs);
                }

                yield return reg;
            }
        }

        static ParameterConversionConfiguration.ReturnConversion CreateReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert, bool handleSubTypes = false)
        {
            return CreateReturnConversion<TFrom>((unusedReturnType, unusedAttributes) => convert, null, handleSubTypes);
        }

        static ParameterConversionConfiguration.ReturnConversion CreateReturnConversion<TFrom>(Func<Type, ExcelReturn, LambdaExpression> returnConversion, Type targetTypeOrNull = null, bool handleSubTypes = false)
        {
            return new ParameterConversionConfiguration.ReturnConversion(returnConversion, targetTypeOrNull, handleSubTypes);
        }
    }
}
