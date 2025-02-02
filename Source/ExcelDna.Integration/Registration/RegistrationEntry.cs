using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    public class RegistrationEntry
    {
        public LambdaExpression FunctionLambda { get; set; }                     // Function which will be registered, and invoked by Excel
        public ExcelFunctionAttribute FunctionAttribute { get; set; }            // ExcelFunctionAttribute which is registered with Excel. Must not be null
        public List<ExcelArgumentAttribute> ArgumentAttributes { get; set; }     // A list of ExcelArgumentAttributes with length equal to the number of parameters in Delegate
        public MethodInfo MethodInfo { get; private set; }                       // The method this entry was originally constructed with (may be useful for transformations).

        // NOTE: 16 parameter max for Expression.GetDelegateType
        public RegistrationEntry(MethodInfo methodInfo)
        {
            MethodInfo = methodInfo;

            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToArray();
            FunctionLambda = Expression.Lambda(Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs);

            // Need to make sure we have explicit 
            FunctionAttribute = methodInfo.GetCustomAttribute<ExcelFunctionAttribute>();
            if (FunctionAttribute == null)
                FunctionAttribute = new ExcelFunctionAttribute { Name = methodInfo.Name };
            else if (string.IsNullOrEmpty(FunctionAttribute.Name))
                FunctionAttribute.Name = methodInfo.Name;

            ArgumentAttributes = new List<ExcelArgumentAttribute>();
            foreach (var pi in methodInfo.GetParameters())
            {
                var argAtt = pi.GetCustomAttribute<ExcelArgumentAttribute>();
                if (argAtt == null)
                    argAtt = new ExcelArgumentAttribute { Name = pi.Name };
                else if (string.IsNullOrEmpty(argAtt.Name))
                    argAtt.Name = pi.Name;

                ArgumentAttributes.Add(argAtt);
            }

            // Special check for final Params argument - transform to an ExcelParamsArgumentAttribute
            // NOTE: This won't work with a custom derived attribute...
            var lastParam = methodInfo.GetParameters().LastOrDefault();
            if (lastParam != null && lastParam.GetCustomAttribute<ParamArrayAttribute>() != null)
            {
                var excelParamsAtt = new Registration.ExcelParamsArgumentAttribute(ArgumentAttributes.Last());
                ArgumentAttributes[ArgumentAttributes.Count - 1] = excelParamsAtt;
            }
        }
    }
}
