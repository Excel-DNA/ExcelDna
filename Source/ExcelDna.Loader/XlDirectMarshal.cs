//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq.Expressions;
using System.Linq;

namespace ExcelDna.Loader
{
    static class XlDirectMarshal
    {
        // Not cleaning up - we don't expect this to use a lot of memory, and think the pool of Excel calculation threads is stable
        readonly static ThreadLocal<XlMarshalContext> MarshalContext = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetMarshalContext() => MarshalContext.Value;

        // This method is only called via AutoFree (for an instance ?)
        // We assume it runs on the same thread as the function call that returned the free-bit value
        public static void FreeMemory()
        {
            GetMarshalContext().FreeMemory();
        }

        // NOTE: This is called in parallel, from a ThreadPool thread
        public static void SetDelegateAndFunctionPointer(XlMethodInfo methodInfo)
        {
            var xlDelegate = GetNativeDelegate(methodInfo);
            methodInfo.DelegateHandle = GCHandle.Alloc(xlDelegate);
            methodInfo.FunctionPointer = Marshal.GetFunctionPointerForDelegate(xlDelegate);
        }

        // NOTE: This is called in parallel, from a ThreadPool thread
        static Delegate GetNativeDelegate(XlMethodInfo methodInfo)
        {
            // We convert
            //
            // double MyFancyFunc(object input1, string input2) { ... }
            //
            // To
            //
            // IntPtr MyFancyFuncWrapped(IntPtr<XlOper12> xlinput0, IntPtr<XlString12 xlinput1)
            // {
            //
            //    XlMarshalContext ctx = GetContextForThisThread();
            //    try
            //    {
            //  
            //      input0 = Convert1(xlinput0);
            //      input1 = Convert2(xlinput1);
            //
            //      result = MyFancyFunc(input0, input1);
            //      xlresult = ctx.ConvertRet(ctx, result),
            //      return xlresult;
            //    }
            //    catch(Exception ex)
            //    {
            //       resultx = HandleEx(ex);
            //       xlresultx = ctx.ConvertRet(resultex);
            //    }
            //
            //    return resultx
            //
            //
            // }

            // Create the new parameters and return value for the wrapper
            // TODO/DM: Consolidate to a single select
            var outerParams = methodInfo.Parameters.Select(p => Expression.Parameter(typeof(IntPtr), p.Name)).ToArray();
            var innerParamExprs = new Expression[outerParams.Length];
            ParameterExpression asyncHandle = null;
            Expression asyncHandleAssign = null;

            for (int i = 0; i < methodInfo.Parameters.Length; i++)
            {
                var pi = methodInfo.Parameters[i];
                if (pi.XlMarshalConvert == null)
                {
                    innerParamExprs[i] = outerParams[i];
                }
                else
                {
                    innerParamExprs[i] = Expression.Call(pi.XlMarshalConvert, outerParams[i]);
                    if (pi.XlType == XlTypes.AsyncHandle)
                    {
                        // We insert an additional cast from the conversion's object return type to the ExcelAsyncHandle type
                        // - we don't have the handle type (defined in ExcelDna.Integration) available when we build ExcelDna.Loader
                        asyncHandle = Expression.Variable(IntegrationMarshalHelpers.ExcelAsyncHandleType, "asyncHandle");
                        asyncHandleAssign = Expression.Assign(asyncHandle, Expression.TypeAs(innerParamExprs[i], IntegrationMarshalHelpers.ExcelAsyncHandleType));
                        innerParamExprs[i] = asyncHandle;
                    }
                }
            }

            Expression innerCall;
            if (methodInfo.MethodInfo != null) // Method and optional Target
            {
                innerCall = Expression.Call(methodInfo.Target == null ? null : Expression.Constant(methodInfo.Target), methodInfo.MethodInfo, innerParamExprs);
            }
            else // LambdaExpression
            {
                innerCall = Expression.Invoke(methodInfo.LambdaExpression, innerParamExprs);
            }
            if (methodInfo.IsExcelAsyncFunction)
            {
                innerCall = Expression.Block(
                    asyncHandleAssign,
                    innerCall);
            }

            // variable to hold XlMarshalContext
            var ctx = Expression.Variable(typeof(XlMarshalContext), "xlMarshalContext");
            var assignCtx = Expression.Assign(ctx, Expression.Call(typeof(XlDirectMarshal), nameof(XlDirectMarshal.GetMarshalContext), null));
            if (methodInfo.HasReturnType)
            {
                var result = Expression.Variable(typeof(IntPtr), "returnValue");
                var resultExpr = result; // Overwrite with conversion if applicable
                if (methodInfo.ReturnType.XlMarshalConvert != null)
                {
                    innerCall = Expression.Call(ctx, methodInfo.ReturnType.XlMarshalConvert, innerCall);
                }
            }

            // Prepare the ex(ception) local variable (TODO/DM why a parameter?)
            var ex = Expression.Variable(typeof(Exception), "ex");
            Expression catchExpression;
            if (methodInfo.IsExcelAsyncFunction)
            {
                // HandleUnhandledException is called by ExcelAsyncHandle.SetException
                var setExceptionMethod = IntegrationMarshalHelpers.ExcelAsyncHandleType.GetMethod("SetException");
                // Need to get instance from parameter list
                catchExpression = Expression.Block(
                            Expression.Call(asyncHandle, setExceptionMethod, ex), // Ignore bool return !?
                            Expression.Empty());
            }
            else
            {
                var exHandler = Expression.Call(typeof(IntegrationHelpers).GetMethod(nameof(IntegrationHelpers.HandleUnhandledException)), ex);
                if (methodInfo.HasReturnType)
                {
                    if (methodInfo.ReturnType.XlType == XlTypes.Xloper)
                    {
                        // We return whatever the result is from the unhandled exception handler
                        catchExpression = Expression.Call(ctx, XlMarshalConversions.ObjectReturn, exHandler);
                    }
                    else
                    {
                        // We return #NUM!, which is better than crashing
                        catchExpression = Expression.Block(exHandler, Expression.Constant(IntPtr.Zero));
                    }
                }
                else
                {
                    catchExpression = Expression.Block(exHandler, Expression.Empty());
                }
            }

            Type delegateType;
            Expression body;
            if (methodInfo.HasReturnType)
            {
                delegateType = XlDirectMarshalTypes.XlFuncs[methodInfo.Parameters.Length];
                body = Expression.Block(
                    typeof(IntPtr),
                    new ParameterExpression[] { ctx },
                    assignCtx,
                    Expression.TryCatch(
                        innerCall,
                        Expression.Catch(ex, catchExpression)));
            }
            else
            {
                delegateType = XlDirectMarshalTypes.XlActs[methodInfo.Parameters.Length];
                if (methodInfo.IsExcelAsyncFunction)
                {
                    body = Expression.Block(
                        new ParameterExpression[] { asyncHandle },
                        Expression.TryCatch(
                            innerCall,
                            Expression.Catch(ex, catchExpression)));
                }
                else
                {
                    body = Expression.TryCatch(
                            innerCall,
                            Expression.Catch(ex, catchExpression));
                }
            }

            return Expression.Lambda(delegateType, body, methodInfo.Name, outerParams).Compile();
        }
    }
}
