using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelDna.Loader
{
    // I'm keeping this call in a separate type, hoping that we only try to resolve the ExcelDna.Integration when we make the InitializeIntegration call
    static class IntegrationLoader
    {
        // This version must match the version declared in ExcelDna.Integration.ExcelIntegration
        const int ExcelIntegrationVersion = 11;
        public static void LoadIntegration()
        {
            ExcelIntegration.ConfigureHost(new IntegrationHost());

            // Check the version declared in the ExcelIntegration class
            int integrationVersion = ExcelIntegration.GetExcelIntegrationVersion();
            if (integrationVersion != ExcelIntegrationVersion)
            {
                // This is not the version we are expecting!
                throw new InvalidOperationException("Invalid ExcelIntegration version detected.");
            }
        }
    }

    class IntegrationHost : IIntegrationHost
    {
        public XlCall.XlReturn TryExcelImpl(int xlFunction, out object result, params object[] parameters)
            => XlCallImpl.TryExcelImpl(xlFunction, out result, parameters);

        public byte[] GetResourceBytes(string resourceName, int type)
            => XlAddIn.GetResourceBytes(resourceName, type);

        public void RegisterMethods(List<MethodInfo> methods)
            => XlRegistration.RegisterMethods(methods);

        public void RegisterMethodsWithAttributes(List<MethodInfo> methods, List<object> functionAttributes, List<List<object>> argumentAttributes)
            => XlRegistration.RegisterMethodsWithAttributes(methods, functionAttributes, argumentAttributes);

        public void RegisterDelegatesWithAttributes(List<Delegate> delegates, List<object> functionAttributes, List<List<object>> argumentAttributes)
            => XlRegistration.RegisterDelegatesWithAttributes(delegates, functionAttributes, argumentAttributes);

        public void RegisterLambdaExpressionsWithAttributes(List<LambdaExpression> lambdaExpressions, List<object> functionAttributes, List<List<object>> argumentAttributes)
            => XlRegistration.RegisterLambdaExpressionsWithAttributes(lambdaExpressions, functionAttributes, argumentAttributes);

        public void RegisterRtdWrapper(string progId, object rtdWrapperOptions, object functionAttribute, List<object> argumentAttributes)
            => XlRegistration.RegisterRtdWrapper(progId, rtdWrapperOptions, functionAttribute, argumentAttributes);
    }
}
