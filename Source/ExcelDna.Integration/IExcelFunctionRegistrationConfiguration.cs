using ExcelDna.Registration;
using System;
using System.Linq.Expressions;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionRegistrationConfiguration
    {
        LambdaExpression GetParameterConversion(Type initialParamType, ExcelParameterRegistration paramRegistration);
        LambdaExpression GetReturnConversion(Type initialReturnType, IExcelFunctionReturn returnRegistration);
    }
}
