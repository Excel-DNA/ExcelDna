using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelParameterConversion
    {
        public MethodInfo MethodInfo { get; private set; }

        public ExcelParameterConversion(MethodInfo methodInfo)
        {
            this.MethodInfo = methodInfo;
        }

        public Func<Type, ExcelParameter, LambdaExpression> GetConversion()
        {
            return (type, paramReg) => CreateConversion(type, paramReg);
        }

        private LambdaExpression CreateConversion(Type type, ExcelParameter paramReg)
        {
            if (type != MethodInfo.ReturnType)
                return null;

            var paramExprs = MethodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            return Expression.Lambda(Expression.Call(MethodInfo, paramExprs), MethodInfo.Name, paramExprs);
        }
    }
}
