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

        public static Expression<Func<object, Microsoft.Office.Interop.Excel.Range>> GetRangeParameterConversion(Type paramType, ExcelParameter paramRegistration)
        {
            if (!IsRange(paramType))
                return null;

            return (object input) => ReferenceToRange(input);
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

        static Microsoft.Office.Interop.Excel.Range ReferenceToRange(object xlInput)
        {
            ExcelReference reference = (ExcelReference)xlInput;
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            int index = sheetName.LastIndexOf("]");
            sheetName = sheetName.Substring(index + 1);
            Microsoft.Office.Interop.Excel.Worksheet ws = app.Sheets[sheetName];
            Microsoft.Office.Interop.Excel.Range target = app.Range[ws.Cells[reference.RowFirst + 1, reference.ColumnFirst + 1], ws.Cells[reference.RowLast + 1, reference.ColumnLast + 1]];

            for (int iInnerRef = 1; iInnerRef < reference.InnerReferences.Count; iInnerRef++)
            {
                ExcelReference innerRef = reference.InnerReferences[iInnerRef];
                Microsoft.Office.Interop.Excel.Range innerTarget = app.Range[ws.Cells[innerRef.RowFirst + 1, innerRef.ColumnFirst + 1], ws.Cells[innerRef.RowLast + 1, innerRef.ColumnLast + 1]];
                target = app.Union(target, innerTarget);
            }

            return target;
        }

        static bool IsRange(Type type)
        {
            return type.IsEquivalentTo(typeof(Microsoft.Office.Interop.Excel.Range));
        }
    }
}
