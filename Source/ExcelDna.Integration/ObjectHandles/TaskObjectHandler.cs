using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class TaskObjectHandler
    {
        public static bool IsUserType(Type t)
        {
            return !AssemblyLoader.IsPrimitiveParameterType(t);
        }

        public static Type ReturnType()
        {
            return typeof(string);
        }

#if AOT_COMPATIBLE
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL3050:RequiresDynamicCode", Justification = "Passes all tests")]
#endif
        public static LambdaExpression ProcessTaskObject(LambdaExpression functionLambda)
        {
            var createHandleMethod = typeof(TaskObjectHandler).GetMethod(nameof(CreateTaskHandle), BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(functionLambda.ReturnType.GetGenericArguments()[0]);
            return ProcessMethod(functionLambda, createHandleMethod);
        }

        public static LambdaExpression ProcessObject(LambdaExpression functionLambda)
        {
            var createHandleMethod = typeof(TaskObjectHandler).GetMethod(nameof(CreateHandle), BindingFlags.Static | BindingFlags.NonPublic);
            return ProcessMethod(functionLambda, createHandleMethod);
        }

#if AOT_COMPATIBLE
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL3050:RequiresDynamicCode", Justification = "Passes all tests")]
#endif
        private static LambdaExpression ProcessMethod(LambdaExpression functionLambda, MethodInfo createHandleMethod)
        {
            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);

            var innerLambda = Expression.Invoke(functionLambda, newParams);
            var callCreateHandle = Expression.Call(createHandleMethod, innerLambda);
            return Expression.Lambda(callCreateHandle, newParams);
        }

        private static async Task<string> CreateTaskHandle<T>(Task<T> data)
        {
            object o = await data;
            return CreateHandle(o);
        }

        private static string CreateHandle(object o)
        {
            return ObjectHandler.GetHandle(o.GetType().ToString(), o);
        }
    }
}
