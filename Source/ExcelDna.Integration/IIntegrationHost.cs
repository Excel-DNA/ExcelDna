using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDna.Integration
{
    interface IIntegrationHost
    {
        XlCall.XlReturn TryExcelImpl(int xlFunction, out object result, params object[] parameters);
        byte[] GetResourceBytes(string resourceName, int type); // types: 0 - Assembly, 1 - Dna file, 2 - Image
        void RegisterMethods(List<MethodInfo> methods);
        void RegisterMethodsWithAttributes(List<MethodInfo> methods, List<object> functionAttributes, List<List<object>> argumentAttributes);
        void RegisterDelegatesWithAttributes(List<Delegate> delegates, List<object> functionAttributes, List<List<object>> argumentAttributes);
        void RegisterLambdaExpressionsWithAttributes(List<LambdaExpression> lambdaExpressions, List<object> functionAttributes, List<List<object>> argumentAttributes);
        void RegisterRtdWrapper(string progId, object rtdWrapperOptions, object functionAttribute, List<object> argumentAttributes);
    }
}
