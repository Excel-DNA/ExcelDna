using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class RangeConversion
    {
        public static IEnumerable<ExcelFunction> UpdateRegistrationsForRangeParameters(this IEnumerable<ExcelFunction> registrations)
        {
            return registrations.Select(UpdateAttributesForRangeParameters);
        }

        public static Expression<Func<object, Microsoft.Office.Interop.Excel.Range>> GetRangeParameterConversion(Type paramType, IExcelFunctionParameter paramRegistration)
        {
            if (!IsRange(paramType))
                return null;

            return (object input) => ExcelConversionUtil.ReferenceToRange((ExcelReference)input);
        }

        static ExcelFunction UpdateAttributesForRangeParameters(ExcelFunction reg)
        {
            var rangeParams = from parWithIndex in reg.FunctionLambda.Parameters.Select((par, i) => new { Parameter = par, Index = i })
                              where IsRange(parWithIndex.Parameter.Type)
                              select parWithIndex;

            foreach (var param in rangeParams)
                reg.ParameterRegistrations[param.Index].ArgumentAttribute.AllowReference = true;

            return reg;
        }

        static bool IsRange(Type type)
        {
            return type.IsEquivalentTo(typeof(Microsoft.Office.Interop.Excel.Range));
        }
    }
}
