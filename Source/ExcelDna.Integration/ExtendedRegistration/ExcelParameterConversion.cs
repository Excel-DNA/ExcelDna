using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelParameterConversion
    {
        private MethodInfo methodInfo;

        public ExcelParameterConversion(MethodInfo methodInfo)
        {
            this.methodInfo = methodInfo;
        }

        public Func<Type, ExcelParameter, LambdaExpression> GetConversion()
        {
            return (type, paramReg) => CreateConversion(type, paramReg);
        }

        private LambdaExpression CreateConversion(Type type, ExcelParameter paramReg)
        {
            if (type != methodInfo.ReturnType)
                return null;

            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            return Expression.Lambda(Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs);
        }
    }
}
