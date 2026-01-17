using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelReturnConversion
    {
        public MethodInfo MethodInfo { get; private set; }

        public ExcelReturnConversion(MethodInfo methodInfo)
        {
            this.MethodInfo = methodInfo;
        }

        public Func<Type, IExcelFunctionReturn, LambdaExpression> GetConversion()
        {
            return (type, returnReg) => CreateConversion(type, returnReg);
        }

#if AOT_COMPATIBLE
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL3050:RequiresDynamicCode", Justification = "Passes all tests")]
#endif
        private LambdaExpression CreateConversion(Type type, IExcelFunctionReturn returnReg)
        {
            ParameterInfo[] parameters = MethodInfo.GetParameters();

            if (parameters.Length != 1 || type != parameters[0].ParameterType)
                return null;

            var paramExprs = parameters
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            return Expression.Lambda(Expression.Call(MethodInfo, paramExprs), MethodInfo.Name, paramExprs);
        }
    }
}
