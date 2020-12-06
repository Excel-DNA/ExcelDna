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

        // Given a delegate and information about the intended export, we instantiate and return the exportable delegate
        // The delegate captures this (singleton) object and then reads the MarshalContext.Value from the ThreadLocal

        // This is an alternative path to XlMethodInfo.CreateDelegateAndFunctionPointer
        public static void SetDelegateAndFunctionPointer(XlMethodInfo methodInfo)
        {
            var delegateType = GetNativeDelegateType(methodInfo);
            var xlDelegate = GetNativeDelegate(methodInfo, delegateType); // Remember _target? ?????
            methodInfo.DelegateHandle = GCHandle.Alloc(xlDelegate);
            methodInfo.FunctionPointer = Marshal.GetFunctionPointerForDelegate(xlDelegate);
        }

        static Delegate GetNativeDelegate(XlMethodInfo methodInfo, Type delegateType)
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
            var innerParamExprs = outerParams.Cast<Expression>().ToArray();  // clone as default - overwrite with conversions where applicable 

            for (int i = 0; i < methodInfo.Parameters.Length; i++)
            {
                var pi = methodInfo.Parameters[i];
                if (pi.DirectMarshalConvert != null)
                {
                    innerParamExprs[i] = Expression.Call(pi.DirectMarshalConvert, innerParamExprs[i]);
                    if (pi.DirectMarshalXlType == XlTypes.AsyncHandle)
                    {
                        // We insert an additional cast from the conversion's object return type to the ExcelAsyncHandle type
                        // - we don't have the handle type (defined in ExcelDna.Integration) available when we build ExcelDna.Loader
                        innerParamExprs[i] = Expression.TypeAs(innerParamExprs[i], IntegrationMarshalHelpers.ExcelAsyncHandleType);
                    }
                }
            }

            // variable to hold XlMarshalContext
            var ctx = Expression.Variable(typeof(XlMarshalContext), "xlMarshalContext");
            var getCtx = Expression.Call(typeof(XlDirectMarshal), nameof(XlDirectMarshal.GetMarshalContext), null);
            var assignCtx = Expression.Assign(ctx, getCtx);
            var wrappingCall = Expression.Call(methodInfo.GetMethodInfo(), innerParamExprs);  // Maybe make more flexible options for XlMethodInfo to be created, e.g. Expressions
            if (methodInfo.ReturnType != null)
            {
                var result = Expression.Variable(typeof(IntPtr), "returnValue");
                var resultExpr = result; // Overwrite with conversion if applicable
                if (methodInfo.ReturnType.DirectMarshalConvert != null)
                {
                    wrappingCall = Expression.Call(ctx, methodInfo.ReturnType.DirectMarshalConvert, wrappingCall);
                }
            }

            BlockExpression block;
            if (methodInfo.ReturnType != null)
            {
                block = Expression.Block(
                    typeof(IntPtr), // Not sure we need this !?
                    new ParameterExpression[] { ctx },
                    assignCtx,
                    wrappingCall);
            }
            else
            {
                block = Expression.Block(
                    new ParameterExpression[] { ctx },
                    assignCtx,
                    wrappingCall);
            }

            var lambda = Expression.Lambda(delegateType, block, methodInfo.Name, outerParams);
            return lambda.Compile();
        }

        static Type GetNativeDelegateType(XlMethodInfo methodInfo)
        {
            if (methodInfo.ReturnType == null)
                return XlDirectMarshalTypes.XlActs[methodInfo.Parameters.Length];
            else
                return XlDirectMarshalTypes.XlFuncs[methodInfo.Parameters.Length];
        }
    }
}
