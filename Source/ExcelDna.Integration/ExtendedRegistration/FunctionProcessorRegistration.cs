using System.Collections.Generic;
using System.Linq;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class FunctionProcessorRegistration
    {
        public static IEnumerable<ExcelDna.Registration.ExcelFunctionRegistration> ProcessFunctionProcessors(this IEnumerable<ExcelDna.Registration.ExcelFunctionRegistration> registrations, IEnumerable<ExcelFunctionProcessor> excelFunctionProcessors, ParameterConversionConfiguration conversionConfig)
        {
            IEnumerable<IExcelFunctionInfo> result = registrations;
            ExcelFunctionRegistrationConfiguration config = new ExcelFunctionRegistrationConfiguration(conversionConfig);
            foreach (ExcelFunctionProcessor p in excelFunctionProcessors.OrderBy(i => i.Name))
            {
                result = p.Invoke(result, config);
            }

            return result.Cast<ExcelDna.Registration.ExcelFunctionRegistration>();
        }
    }
}
