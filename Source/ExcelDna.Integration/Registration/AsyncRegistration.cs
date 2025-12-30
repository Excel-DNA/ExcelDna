using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.ExtendedRegistration;
using ExcelDna.Integration.ObjectHandles;

#if USE_WINDOWS_FORMS
using ExcelDna.Logging;
#else
using ExcelDna.Integration.Win32;
#endif

namespace ExcelDna.Registration
{
    public static class AsyncRegistration
    {
        /// <summary>
        /// Wraps methods that are marked with [ExcelFunction] and:
        /// * Return IObservable, or
        /// * Return Task or 
        /// * Are marked with [ExcelAsyncFunction] attribute.
        /// 
        /// If the function takes as last parameter a CancellationToken, this will be hooked up to the async function cancellation.
        /// </summary>
        /// <remarks>NOTE: Currently supports functions with no more than sixteen parameters (fifteen for the native async),
        ///               due to the limitation of the Expression.LambdaExpression call we use.
        /// </remarks>
        /// <param name="registrations">The list is RegistrationEntries to process.</param>
        /// <param name="nativeAsyncIfAvailable">Under Excel 2010 and later, indicates whether the native async feature should be used.
        /// Does not apply to functions returning IObservable.</param>
        /// <returns>The list of RegistrationEntries, with the affected functions wrapped to allow registration in Excel.</returns>
        public static IEnumerable<ExcelDna.Registration.ExcelFunctionRegistration> ProcessAsyncRegistrations(this IEnumerable<ExcelDna.Registration.ExcelFunctionRegistration> registrations, bool nativeAsyncIfAvailable = false)
        {
            // Decide whether Tasks should be using native async
            bool useNativeAsync = nativeAsyncIfAvailable && ExcelDnaUtil.ExcelVersion >= 14.0;
            if (useNativeAsync) NativeAsyncTaskUtil.Initialize();

            foreach (var reg in registrations)
            {
                try
                {
                    if (ReturnsObservable(reg.FunctionLambda))
                    {
                        ParameterConversionRegistration.ApplyParameterConversions(reg, ObjectHandleRegistration.GetParameterConversionConfiguration());
                        reg.FunctionLambda = WrapMethodObservable(reg.FunctionLambda, reg.Return.CustomAttributes);
                    }
                    else if (ReturnsTask(reg.FunctionLambda) || reg.FunctionAttribute is ExcelAsyncFunctionAttribute)
                    {
                        ParameterConversionRegistration.ApplyParameterConversions(reg, ObjectHandleRegistration.GetParameterConversionConfiguration());
                        if (HasCancellationToken(reg.FunctionLambda))
                        {
                            reg.FunctionLambda = useNativeAsync ? WrapMethodNativeAsyncTaskWithCancellation(reg.FunctionLambda)
                                                                : WrapMethodRunTaskWithCancellation(reg.FunctionLambda, reg.Return.CustomAttributes);
                            // Also need to strip out the info for the last argument which is the CancellationToken
                            reg.ParameterRegistrations.RemoveAt(reg.ParameterRegistrations.Count - 1);
                        }
                        else
                        {
                            reg.FunctionLambda = useNativeAsync ? WrapMethodNativeAsyncTask(reg.FunctionLambda)
                                                                : WrapMethodRunTask(reg.FunctionLambda, reg.Return.CustomAttributes);
                        }

                        if (reg.FunctionAttribute is ExcelAsyncFunctionAttribute)
                            reg.FunctionAttribute = new ExcelFunctionAttribute(reg.FunctionAttribute);
                    }
                    // else do nothing to this registration
                }
                catch (Exception ex)
                {
                    LogDisplay.WriteLine("Exception while registering method {0} - {1}", reg.FunctionAttribute.Name, ex.ToString());
                    continue;
                }

                yield return reg;
            }
        }

        static bool ReturnsObservable(LambdaExpression functionLambda)
        {
            return functionLambda.ReturnType.IsGenericType && functionLambda.ReturnType.GetGenericTypeDefinition() == typeof(IObservable<>);
        }

        static bool ReturnsTask(LambdaExpression functionLambda)
        {
            return functionLambda.ReturnType.IsGenericType && functionLambda.ReturnType.GetGenericTypeDefinition() == typeof(Task<>);
        }

        static bool HasCancellationToken(LambdaExpression functionLambda)
        {
            var pis = functionLambda.Parameters;
            return pis.Any() && pis.Last().Type == typeof(CancellationToken);
        }

        static LambdaExpression WrapMethodRunTask(LambdaExpression functionLambda, List<object> returnCustomAttributes)
        {
            /* Either, from a lambda expression wrapping a method that looks like this:
             * 
             *      static Task<string> myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static object myFunc(string name, int msDelay)
             *      {
             *          return AsyncTaskUtil.RunTask<string>(
             *              "myFunc:XXX", 
             *              new object[] {(object)name, (object)msDelay}, 
             *              () => myFunc(name, msDelay));
             *      }
             * 
             * Or, from a lambda expression wrapping a method that looks like this (not returning a Task):
             * 
             *      static string myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this (with RunAsTaskXXX):
             * 
             *      static object myFunc(string name, int msDelay)
             *      {
             *          return AsyncTaskUtil.RunAsTask<string>(
             *              "myFunc:XXX", 
             *              new object[] {(object)name, (object)msDelay}, 
             *              () => myFunc(name, msDelay));
             *      }
             */

            bool returnsTask = ReturnsTask(functionLambda);
            Type returnType = returnsTask ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            bool userType = TaskObjectHandler.IsUserType(returnType);
            if (userType)
            {
                if (!ObjectHandleRegistration.HasExcelHandle(returnCustomAttributes))
                    throw new Exception($"Unsupported task return type {returnType}.");

                ObjectHandleRegistration.ClearExcelHandle(returnCustomAttributes);
            }

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = returnsTask ? "RunTask" : "RunAsTask";
            if (userType)
                runMethodName = runMethodName.Replace("Task", "TaskObject");

            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = returnsTask ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            if (userType)
                newReturnType = TaskObjectHandler.ReturnType();

            // Build up the RunTaskWithC... method with the right generic type argument
#pragma warning disable IL2060 // Guaranteed to work by the SourceGenerator adding to methodRefs.
            var runMethod = typeof(ExcelAsyncUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);
#pragma warning restore IL2060

            // Get the function name
            var nameExp = Expression.Constant(functionLambda.Name + ":" + Guid.NewGuid().ToString("N"));

            // Make the new params for the wrapper - they look exactly like the functionLambda's parameters
            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Also cast params to Object and put into a fresh object[] array for the RunTask call
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);
            LambdaExpression adaptedFunctionLambda = userType ?
                (returnsTask ? TaskObjectHandler.ProcessTaskObject(functionLambda) : TaskObjectHandler.ProcessObject(functionLambda)) :
                functionLambda;
            var innerLambda = Expression.Lambda(Expression.Invoke(adaptedFunctionLambda, newParams));

            // This is the call to RunTask, taking the name, param array and the (capturing) lambda (called with no arguments)
            var callTaskRun = Expression.Call(runMethod, nameExp, paramsArrayExp, innerLambda);

            // Wrap with all the parameters
            return Expression.Lambda(callTaskRun, functionLambda.Name, newParams);
        }

        static LambdaExpression WrapMethodRunTaskWithCancellation(LambdaExpression functionLambda, List<object> returnCustomAttributes)
        {
            /* Either, from a lambda expression that looks like this:
             * 
             *      static Task<string> myFuncWithCancel(string name, int msDelay, CancellationToken ct) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static object crDelayedHello(string name, int msDelay)
             *      {
             *          return AsyncTaskUtil.RunTaskWithCancellation<string>(
             *              "myFuncWithCancel:XXX", 
             *              new object[] {(object)name, (object)msDelay}, 
             *              (ct) => myFuncWithCancel(name, msDelay, ct));
             *      }
             * 
             * Or, from a lambda expression that looks like this (not returning a Task):
             * 
             *      static string myFuncWithCancel(string name, int msDelay, CancellationToken ct) {...}
             * 
             *   we create a lambda expression that looks like this (with RunAsTaskXXX):
             * 
             *      object crDelayedHello(string name, int msDelay)
             *      {
             *          return AsyncTaskUtil.RunAsTaskWithCancellation<string>(
             *              "myFuncWithCancel:XXX", 
             *              new object[] {(object)name, (object)msDelay}, 
             *              (ct) => myFuncWithCancel(name, msDelay, ct));
             *      }
             */

            bool returnsTask = ReturnsTask(functionLambda);
            Type returnType = returnsTask ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            bool userType = TaskObjectHandler.IsUserType(returnType);
            if (userType)
            {
                if (!ObjectHandleRegistration.HasExcelHandle(returnCustomAttributes))
                    throw new Exception($"Unsupported task return type {returnType}.");

                ObjectHandleRegistration.ClearExcelHandle(returnCustomAttributes);
            }

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = returnsTask ? "RunTaskWithCancellation" : "RunAsTaskWithCancellation";
            if (userType)
                runMethodName = runMethodName.Replace("Task", "TaskObject");

            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = returnsTask ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            if (userType)
                newReturnType = TaskObjectHandler.ReturnType();

            // Build up the RunTaskWithC... method with the right generic type argument
#pragma warning disable IL2060 // Guaranteed to work by the SourceGenerator adding to methodRefs.
            var runMethod = typeof(ExcelAsyncUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);
#pragma warning restore IL2060

            // Get the function name - passed as the first argument to RunTask...
            var nameExp = Expression.Constant(functionLambda.Name + ":" + Guid.NewGuid().ToString("N"));

            // ... and parameters excluding that CancellationToken
            //     (for the exported lambda and captured for the inner Lambda)
            var newParams = functionLambda.Parameters
                            .Where(param => param.Type != typeof(CancellationToken))
                            .Select(param => Expression.Parameter(param.Type, param.Name))
                            .ToList();

            // Also cast params to Object and put into a fresh object[] array for the second argument to RunTask...
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);

            // Now add the extra CancellationToken parameter
            var ctParamExp = Expression.Parameter(typeof(CancellationToken));
            var allParams = new List<ParameterExpression>(newParams) { ctParamExp };
            LambdaExpression adaptedFunctionLambda = userType ?
                (returnsTask ? TaskObjectHandler.ProcessTaskObject(functionLambda) : TaskObjectHandler.ProcessObject(functionLambda)) :
                functionLambda;
            var innerLambda = Expression.Lambda(Expression.Invoke(adaptedFunctionLambda, allParams), ctParamExp);

            // This is the call to RunTask, taking the name, param array and the (capturing) lambda
            var callTaskRun = Expression.Call(runMethod, nameExp, paramsArrayExp, innerLambda);

            // Wrap with all the parameters, and Compile to a Delegate
            return Expression.Lambda(callTaskRun, functionLambda.Name, newParams);
        }

        static LambdaExpression WrapMethodNativeAsyncTask(LambdaExpression functionLambda)
        {
#if COM_GENERATED
            throw new NotImplementedException("WrapMethodNativeAsyncTask is not supported in AOT.");
#else
            /* Either, from a lambda expression that looks like this:
             * 
             *      static Task<string> myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static void myFunc(string name, int msDelay, ExcelAsyncHandle asyncHandle)
             *       {
             *           NativeAsyncTaskUtil.RunTask(() => myFunc(name, msDelay), asyncHandle);
             *       }
             *
             * Or, from a lambda expression that looks like this (not returning a Task):
             * 
             *      static string myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this (with RunAsTaskXXX):
             * 
             *      static void myFunc(string name, int msDelay, ExcelAsyncHandle asyncHandle)
             *       {
             *           NativeAsyncTaskUtil.RunAsTask(() => myFunc(name, msDelay), asyncHandle);
             *       }
             */

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = ReturnsTask(functionLambda) ? "RunTask" : "RunAsTask";
            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = ReturnsTask(functionLambda) ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;

            // Make the new params for the wrapper - they look exactly like the functionLambda's parameters
            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Build up the RunTaskWithC... method with the right generic type argument
            var runMethod = typeof(NativeAsyncTaskUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, newParams));

            // Create the AsyncHandle param
            var asyncHandleParam = Expression.Parameter(typeof(ExcelAsyncHandle), "asyncHandle");
            // This is the call to RunTask, taking taking the (capturing) lambda and async handle
            var callTaskRun = Expression.Call(runMethod, innerLambda, asyncHandleParam);

            // Wrap with all the parameters
            var allParams = new List<ParameterExpression>(newParams) { asyncHandleParam };
            return Expression.Lambda(callTaskRun, functionLambda.Name, allParams);
#endif
        }

        static LambdaExpression WrapMethodNativeAsyncTaskWithCancellation(LambdaExpression functionLambda)
        {
#if COM_GENERATED
            throw new NotImplementedException("WrapMethodNativeAsyncTaskWithCancellation is not supported in AOT.");
#else
            /* Either, from a lambda expression that looks like this:
             * 
             *      static Task<string> myFunc(string name, int msDelay, CancellationToken ct) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static void myFunc(string name, int msDelay, ExcelAsyncHandle asyncHandle)
             *       {
             *           NativeAsyncTaskUtil.RunTaskWithCancellation((ct) => myFunc(name, msDelay, ct), asyncHandle);
             *       }
             *
             * Or, from a lambda expression that looks like this (not returning a Task):
             * 
             *      static string myFunc(string name, int msDelay, CancellationToken ct) {...}
             * 
             *   we create a lambda expression that looks like this (with RunAsTaskXXX):
             * 
             *      static void myFunc(string name, int msDelay, ExcelAsyncHandle asyncHandle)
             *       {
             *           NativeAsyncTaskUtil.RunAsTaskWithCancellation((ct) => myFunc(name, msDelay, ct), asyncHandle);
             *       }
             */

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = ReturnsTask(functionLambda) ? "RunTaskWithCancellation" : "RunAsTaskWithCancellation";
            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = ReturnsTask(functionLambda) ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            // Build up the RunTaskWithC... method with the right generic type argument
            var runMethod = typeof(NativeAsyncTaskUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);

            // ... and parameters excluding the CancellationToken (used both in the call to RunTask and captured for the inner Lambda)
            var newParams = functionLambda.Parameters
                            .Where(param => param.Type != typeof(CancellationToken))
                            .Select(param => Expression.Parameter(param.Type, param.Name))
                            .ToList();

            // Set up the parameters and inner lambda, with call to RunTaskW...
            // Not sure if I can re-use the inner CancellationToken parameter...
            var ctParam = Expression.Parameter(typeof(CancellationToken), "cancellationToken");
            var innerParams = new List<ParameterExpression>(newParams) { ctParam };
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, innerParams), ctParam);

            // Create the AsyncHandle param
            var asyncHandleParam = Expression.Parameter(typeof(ExcelAsyncHandle), "asyncHandle");
            // This is the call to RunTask, taking the name, param array and the (capturing) lambda
            var callTaskRun = Expression.Call(runMethod, innerLambda, asyncHandleParam);

            // Wrap with all the parameters, and Compile to a Delegate
            var allParams = new List<ParameterExpression>(newParams) { asyncHandleParam };
            return Expression.Lambda(callTaskRun, functionLambda.Name, allParams);
#endif
        }

        static LambdaExpression WrapMethodObservable(LambdaExpression functionLambda, List<object> returnCustomAttributes)
        {
            /* Either, from a lambda expression that looks like this:
             * 
             *      static IObservable<string> myFunc(string name, int msDelay) {...}
             * 
             *   we create a lambda expression that looks like this:
             * 
             *      static object myFunc(string name, int msDelay)
             *      {
             *          return ObservableRtdUtil.Observer<string>(          // obsMethod
             *              "myFunc:XXX",                                   // name
             *              new object[] {(object)name, (object)msDelay},   // paramsArrayExp
             *              () => myFunc(name, msDelay));                   // innerLambda
             *      }
             */

            // mi returns some kind of IObservable<T>. What is T? 
            Type returnType = functionLambda.ReturnType.GetGenericArguments()[0];
            bool userType = TaskObjectHandler.IsUserType(returnType);
            if (userType)
            {
                if (!ObjectHandleRegistration.HasExcelHandle(returnCustomAttributes))
                    throw new Exception($"Unsupported observable return type {returnType}.");

                ObjectHandleRegistration.ClearExcelHandle(returnCustomAttributes);
            }

            // Build up the Observe method with the right generic type argument
            var obsMethod = typeof(ExcelAsyncUtil)
                                .GetMember(userType ? "ObserveObject" : "Observe", MemberTypes.Method, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                                .Cast<MethodInfo>().First(i => i.IsGenericMethodDefinition)
                                .MakeGenericMethod(returnType);

            // Get the function name
            var nameExp = Expression.Constant(functionLambda.Name + ":" + Guid.NewGuid().ToString("N"));
            // ... and parameters (used both in the call to Observe and captured for the inner Lambda

            // Make the new params for the wrapper - they look exactly like the functionLambda's parameters
            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Cast params to Object and put into a fresh object[] array for the Observe call
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, newParams));

            // This is the call to Observe, taking the name, param array and the (capturing) lambda
            var callTaskRun = Expression.Call(obsMethod, nameExp, paramsArrayExp, innerLambda);

            // Wrap with all the parameters, and Compile to a Delegate
            return Expression.Lambda(callTaskRun, functionLambda.Name, newParams);
        }
    }
}
