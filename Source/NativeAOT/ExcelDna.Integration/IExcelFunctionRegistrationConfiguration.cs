using System;
using System.Linq.Expressions;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionRegistrationConfiguration
    {
        LambdaExpression GetParameterConversion(Type initialParamType, IExcelFunctionParameter paramRegistration);
        LambdaExpression GetReturnConversion(Type initialReturnType, IExcelFunctionReturn returnRegistration);
    }
}
