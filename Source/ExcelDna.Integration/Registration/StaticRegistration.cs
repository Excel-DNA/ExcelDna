using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelDna.Registration
{
    public static class StaticRegistration
    {
        public interface IDirectMarshalTypeAdapter
        {
            IntPtr GetFunctionPointerForDelegate(Delegate d, int parameters);
            IntPtr GetActionPointerForDelegate(Delegate d, int parameters);

            Type GetFunctionType(int parameters);
            Type GetActionType(int parameters);

            Delegate CreateFunctionDelegate(Func<Delegate> delegateFactory, int parameters);
            Delegate CreateActionDelegate(Func<Delegate> delegateFactory, int parameters);
        }

        public static List<MethodInfo> MethodsForRegistration { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelParameterConversions { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelReturnConversions { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelFunctionExecutionHandlerSelectors { get; } = new List<MethodInfo>();
        public static List<Integration.ITypeHelper> ExcelAddIns { get; } = new List<Integration.ITypeHelper>();
        public static List<object> AssemblyAttributes { get; } = new List<object>();
        public static IDirectMarshalTypeAdapter DirectMarshalTypeAdapter { get; set; }
    }
}
