using System.Collections.Generic;
using System.Reflection;

namespace ExcelDna.Registration
{
    public static class StaticRegistration
    {
        public static List<MethodInfo> MethodsForRegistration { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelParameterConversions { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelReturnConversions { get; } = new List<MethodInfo>();
        public static List<MethodInfo> ExcelFunctionExecutionHandlerSelectors { get; } = new List<MethodInfo>();
        public static List<Integration.ITypeHelper> ExcelAddIns { get; } = new List<Integration.ITypeHelper>();
        public static List<object> AssemblyAttributes { get; } = new List<object>();
    }
}
