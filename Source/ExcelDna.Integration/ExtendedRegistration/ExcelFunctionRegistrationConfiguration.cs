using System;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelFunctionRegistrationConfiguration : IExcelFunctionRegistrationConfiguration
    {
        private ParameterConversionConfiguration conversionConfig;

        public ExcelFunctionRegistrationConfiguration(ParameterConversionConfiguration conversionConfig)
        {
            this.conversionConfig = conversionConfig;
        }

        public LambdaExpression GetParameterConversion(Type initialParamType, IExcelFunctionParameter paramRegistration)
        {
            return ParameterConversionRegistration.GetParameterConversion(conversionConfig, initialParamType, paramRegistration);
        }

        public LambdaExpression GetReturnConversion(Type initialReturnType, IExcelFunctionReturn returnRegistration)
        {
            return ParameterConversionRegistration.GetReturnConversion(conversionConfig, initialReturnType, returnRegistration);
        }
    }
}
