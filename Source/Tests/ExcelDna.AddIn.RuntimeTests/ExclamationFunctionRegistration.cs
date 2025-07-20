using ExcelDna.Integration;
using System.Linq.Expressions;

namespace ExcelDna.AddIn.RuntimeTests
{
    internal static class ExclamationFunctionRegistration
    {
        [ExcelFunctionProcessor]
        public static IEnumerable<IExcelFunctionInfo> ProcessExclamationFunctions(IEnumerable<IExcelFunctionInfo> registrations, IExcelFunctionRegistrationConfiguration config)
        {
            foreach (var reg in registrations)
            {
                if (reg.FunctionAttribute.Name == "MySayHelloWithExclamation")
                {
                    var concatMethod = typeof(string).GetMethod("Concat", new[] { typeof(string), typeof(string) });
                    var newBody = Expression.Call(concatMethod!, reg.FunctionLambda.Body, Expression.Constant("!"));
                    reg.FunctionLambda = Expression.Lambda<Func<string, string>>(newBody, reg.FunctionLambda.Parameters);
                }

                yield return reg;
            }
        }
    }
}
