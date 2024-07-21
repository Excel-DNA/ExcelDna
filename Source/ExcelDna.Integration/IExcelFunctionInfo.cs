using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelDna.Integration
{
    public interface IExcelFunctionInfo
    {
        ExcelFunctionAttribute FunctionAttribute { get; }
        List<IExcelFunctionParameter> Parameters { get; }
        IExcelFunctionReturn Return { get; }
        List<object> CustomAttributes { get; }

        LambdaExpression FunctionLambda { get; set; }
    }
}
