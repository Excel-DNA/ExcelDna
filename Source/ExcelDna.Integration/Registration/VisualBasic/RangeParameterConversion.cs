#if !COM_GENERATED

using System.Linq;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Collections.Generic;
using System;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.Registration.VisualBasic
{
    [Microsoft.VisualBasic.CompilerServices.StandardModule]
    public static class RangeParameterConversion
    {
        public static Range ReferenceToRange(object xlInput)
        {
            ExcelReference reference = (ExcelReference)xlInput;  // Will throw some Exception if not valid, which will be returned as #VALUE
            Application app = (Application)ExcelDnaUtil.Application;

            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            int index = sheetName.LastIndexOf("]");
            sheetName = sheetName.Substring(index + 1);
            Worksheet ws = app.Sheets[sheetName];
            Range target = app.Range[ws.Cells[reference.RowFirst + 1, reference.ColumnFirst + 1],
                                    ws.Cells[reference.RowLast + 1, reference.ColumnLast + 1]];

            for (long iInnerRef = 1; iInnerRef < reference.InnerReferences.Count; iInnerRef++)
            {
                ExcelReference innerRef = reference.InnerReferences[(int)iInnerRef];
                Range innerTarget = app.Range[ws.Cells[innerRef.RowFirst + 1, innerRef.ColumnFirst + 1],
                                            ws.Cells[innerRef.RowLast + 1, innerRef.ColumnLast + 1]];
                target = app.Union(target, innerTarget);
            }
            return target;
        }

        private static ExcelFunctionRegistration UpdateAttributesForRangeParameters(ExcelFunctionRegistration reg)
        {
            var rangeParams = from parWithIndex in reg.FunctionLambda.Parameters.Select((par, i) => new { Parameter = par, Index = i })
                              where parWithIndex.Parameter.Type.IsEquivalentTo(typeof(Range))
                              select parWithIndex;

            foreach (var param in rangeParams)
            {
                reg.ParameterRegistrations[param.Index].ArgumentAttribute.AllowReference = true;
            }

            return reg;
        }

        // Must be run before the parameter conversions
        public static IEnumerable<ExcelFunctionRegistration> UpdateRegistrationsForRangeParameters(this IEnumerable<ExcelFunctionRegistration> regs)
        {
            return regs.Select(UpdateAttributesForRangeParameters);
        }

        // NOTE: This parameter conversion should be registered to run for all types (with 'null' as the TypeFilter)
        // so that the COM-friendly type equivalence check here can be done, instead of exact type check.
        public static Expression<Func<object, Range>> ParameterConversion(Type paramType, ExcelParameterRegistration paramRegistration)
        {
            if (paramType.IsEquivalentTo(typeof(Range)))
            {
                return (object input) => ReferenceToRange(input);
            }
            return null;
        }
    }
}

#endif
