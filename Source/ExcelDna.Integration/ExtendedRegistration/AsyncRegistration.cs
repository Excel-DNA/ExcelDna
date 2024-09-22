using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class AsyncRegistration
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
        public static IEnumerable<ExcelFunction> ProcessAsyncRegistrations(this IEnumerable<ExcelFunction> registrations, bool nativeAsyncIfAvailable = false)
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
                        ParameterConversionRegistration.ApplyParameterConversions(reg, ObjectHandles.ObjectHandleRegistration.GetParameterConversionConfiguration());
                        reg.FunctionLambda = WrapMethodObservable(reg.FunctionLambda);
                    }
                    else if (ReturnsTask(reg.FunctionLambda) || reg.FunctionAttribute is ExcelAsyncFunctionAttribute)
                    {
                        ParameterConversionRegistration.ApplyParameterConversions(reg, ObjectHandles.ObjectHandleRegistration.GetParameterConversionConfiguration());
                        if (HasCancellationToken(reg.FunctionLambda))
                        {
                            reg.FunctionLambda = useNativeAsync ? WrapMethodNativeAsyncTaskWithCancellation(reg.FunctionLambda)
                                                                : WrapMethodRunTaskWithCancellation(reg.FunctionLambda);
                            // Also need to strip out the info for the last argument which is the CancellationToken
                            reg.ParameterRegistrations.RemoveAt(reg.ParameterRegistrations.Count - 1);
                        }
                        else
                        {
                            reg.FunctionLambda = useNativeAsync ? WrapMethodNativeAsyncTask(reg.FunctionLambda)
                                                                : WrapMethodRunTask(reg.FunctionLambda);
                        }
                    }
                    // else do nothing to this registration
                }
                catch (Exception ex)
                {
                    Logging.LogDisplay.WriteLine("Exception while registering method {0} - {1}", reg.FunctionAttribute.Name, ex.ToString());
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

        static LambdaExpression WrapMethodRunTask(LambdaExpression functionLambda)
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

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = ReturnsTask(functionLambda) ? "RunTask" : "RunAsTask";
            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = ReturnsTask(functionLambda) ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            // Build up the RunTaskWithC... method with the right generic type argument
            var runMethod = typeof(ExcelAsyncUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);

            // Get the function name
            var nameExp = Expression.Constant(functionLambda.Name + ":" + Guid.NewGuid().ToString("N"));

            // Make the new params for the wrapper - they look exactly like the functionLambda's parameters
            var newParams = functionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Also cast params to Object and put into a fresh object[] array for the RunTask call
            var paramsArray = newParams.Select(p => Expression.Convert(p, typeof(object)));
            var paramsArrayExp = Expression.NewArrayInit(typeof(object), paramsArray);
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, newParams));

            // This is the call to RunTask, taking the name, param array and the (capturing) lambda (called with no arguments)
            var callTaskRun = Expression.Call(runMethod, nameExp, paramsArrayExp, innerLambda);

            // Wrap with all the parameters
            var lambda = Expression.Lambda(callTaskRun, functionLambda.Name, newParams);
            return lambda;
        }

        static LambdaExpression WrapMethodRunTaskWithCancellation(LambdaExpression functionLambda)
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

            // Either RunTask or RunAsTask, depending on whether the method returns Task<string> or string
            string runMethodName = ReturnsTask(functionLambda) ? "RunTaskWithCancellation" : "RunAsTaskWithCancellation";
            // mi returns some kind of Task<T>. What is T? 
            var newReturnType = ReturnsTask(functionLambda) ? functionLambda.ReturnType.GetGenericArguments()[0] : functionLambda.ReturnType;
            // Build up the RunTaskWithC... method with the right generic type argument
            var runMethod = typeof(ExcelAsyncUtil)
                                .GetMember(runMethodName, MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
                                .Cast<MethodInfo>().First()
                                .MakeGenericMethod(newReturnType);

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
            var innerLambda = Expression.Lambda(Expression.Invoke(functionLambda, allParams), ctParamExp);

            // This is the call to RunTask, taking the name, param array and the (capturing) lambda
            var callTaskRun = Expression.Call(runMethod, nameExp, paramsArrayExp, innerLambda);

            // Wrap with all the parameters, and Compile to a Delegate
            return Expression.Lambda(callTaskRun, functionLambda.Name, newParams);
        }

        static LambdaExpression WrapMethodNativeAsyncTask(LambdaExpression functionLambda)
        {
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
        }

        static LambdaExpression WrapMethodNativeAsyncTaskWithCancellation(LambdaExpression functionLambda)
        {
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
        }

        static LambdaExpression WrapMethodObservable(LambdaExpression functionLambda)
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
            var returnType = functionLambda.ReturnType.GetGenericArguments()[0];
            // Build up the Observe method with the right generic type argument
            var obsMethod = typeof(ExcelAsyncUtil)
                                .GetMember("Observe", MemberTypes.Method, BindingFlags.Static | BindingFlags.Public)
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
