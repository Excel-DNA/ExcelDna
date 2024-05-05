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
            foreach (var reg in registrations)
            {
                if (reg.FunctionLambda.ReturnType == typeof(ExcelObjectHandle))
                {


                    var rc = CreateReturnConversion((ExcelObjectHandle value) => ObjectHandles.Util.ReturnConversionNew(value, reg.FunctionAttribute.Name));
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
