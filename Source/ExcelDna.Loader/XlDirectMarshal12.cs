using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq.Expressions;
using System.Linq;
using System.Reflection;

namespace ExcelDna.Loader
{
    // We have one XlMarshalContext per thread
    // It is never collected
    unsafe class XlMarshalContext
    {
        // TODO/DM: Test again that these are actually meaningful - it seems strange
        readonly static object boxedZero = 0.0;
        readonly static object boxedOne = 1.0;
        //        readonly object excelEmpty = IntegrationMarshalHelpers.GetExcelEmptyValue();

        // These are fixed size, and could be allocated as a single struct or block.
        // Strings of any length, in Xloper or direct, using max length fixed buffer
        XlString12* _pStringBufferReturn;
        double* _pDoubleReturn; // Also used for DateTime
        short* _pBoolReturn;

        // All the in-place Xloper types
        XlOper12* _pXloperReturn;

        // Used for single-element array return, allowing allocation-free return in this case
        XlOper12* _pXloperArraySingletonReturn;

        XlMarshalDoubleArrayContext _rank1DoubleArrayContext;
        XlMarshalDoubleArrayContext _rank2DoubleArrayContext;

        XlMarshalOperArrayContext _rank1OperArrayContext;
        XlMarshalOperArrayContext _rank2OperArrayContext;

        public XlMarshalContext()
        {
            int size;
            // StringReturn
            size = Marshal.SizeOf(typeof(XlString12)) + ((XlString12.MaxLength - 1) /* 1 char is in Data[1] */ * 2 /* 2 bytes per char */);
            _pStringBufferReturn = (XlString12*)Marshal.AllocCoTaskMem(size);

            // DateTimeReturn
            size = Marshal.SizeOf(typeof(double));
            _pDoubleReturn = (double*)Marshal.AllocCoTaskMem(size);

            size = Marshal.SizeOf(typeof(short));
            _pBoolReturn = (short*)Marshal.AllocCoTaskMem(size);

            // XloperReturn
            size = Marshal.SizeOf(typeof(XlOper12));
            _pXloperReturn = (XlOper12*)Marshal.AllocCoTaskMem(size);

            _rank1DoubleArrayContext = new XlMarshalDoubleArrayContext(1);
            _rank2DoubleArrayContext = new XlMarshalDoubleArrayContext(2);

            _rank1OperArrayContext = new XlMarshalOperArrayContext(1, false);
            _rank2OperArrayContext = new XlMarshalOperArrayContext(2, false);
        }

        // RULE: Return conversions must not throw exceptions (they might run in the exception handler)
        // RULE: Param conversions can throw exceptions

        public IntPtr DoubleToXloperReturn(double d)
        {
            _pXloperReturn->numValue = d;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr DoublePtrReturn(double d)
        {
            *_pDoubleReturn = d;
            return (IntPtr)_pDoubleReturn;
        }

        public IntPtr StringReturn(string str)
        {
            // We maintain compatible behaviour with the CustomMarhsalling, which would return null pointers directly (without calling marshalling)
            // DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!
            if (str == null)
                return IntPtr.Zero;

            XlString12* pdest = _pStringBufferReturn;
            ushort charCount = (ushort)Math.Min(str.Length, XlString12.MaxLength);

            // TODO/DM Try to remember why we did this instead of Marshal.Copy
            fixed (char* psrc = str)
            {
                char* ps = psrc;
                char* pd = pdest->Data;
                for (int k = 0; k < charCount; k++)
                {
                    *(pd++) = *(ps++);
                }
            }
            pdest->Length = charCount;

            return (IntPtr)_pStringBufferReturn;
        }

        public IntPtr DateTimeToDoublePtrReturn(DateTime dateTime)
        {
            try
            {
                *_pDoubleReturn = dateTime.ToOADate();
                return (IntPtr)_pDoubleReturn;
            }
            catch
            {
                // This case is where the range of the OADate is exceeded, e.g. a year before 0100.
                // We'd like to return #VALUE, but we're registered as a double*
                // returning IntPtr.Zero will give us #NUM in Excel
                return IntPtr.Zero;
            }
        }

        public IntPtr DateTimeToXloperReturn(DateTime dateTime)
        {
            // TODO/DM DOCUMENT: 
            // In the case where we have a date that cannot be converted to an OleAutomation date, e.g. year < 0100
            // it is not a valid date in Excel (no dates before 1900 are valid)
            // We return #VALUE without calling our internal Exception handler

            try
            {
                _pXloperReturn->numValue = dateTime.ToOADate();
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            catch
            {
                // This is a case where we have a date that cannot be converted to an OleAutomation date, e.g. year < 0100
                // Certainly it is not a valid date in Excel (no dates before 1900 are valid)
                // But we must not crash - return #VALUE instead
                _pXloperReturn->errValue = IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                _pXloperReturn->xlType = XlType12.XlTypeError;
            }
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr ObjectToXloperReturn(object str)
        {
            // We maintain compatible behaviour with the CustomMarhsalling, which would return null pointers directly (without calling marshalling)
            // DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!
            if (str == null)
                return IntPtr.Zero;

            throw new NotImplementedException();
        }

        // Input parameter conversions (also for XlCall.Excel return values) - static, no context
        public static double DoublePtrParam(IntPtr pd)
        {
            return *(double*)pd;
        }

        public static string StringParam(IntPtr pstrValue)
        {
            XlString12* pString = (XlString12*)pstrValue;
            return new string(pString->Data, 0, pString->Length);
        }

        public static DateTime DateTimeFromDoublePtrParam(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the OADate range
            //          This code will raise an exception and we have to handle it or return immediately or something

            double dateSerial = *(double*)pNativeData;
            return DateTime.FromOADate(dateSerial);
        }
    }

    unsafe class XlMarshalDoubleArrayContext
    {
        int _rank;
        XlFP12* _pNative; // For managed -> native returns

        public XlMarshalDoubleArrayContext(int rank)
        {
            _rank = rank;
        }
    }

    unsafe class XlMarshalOperArrayContext
    {
        int _rank;
        // These used for array return
        List<XlMarshalOperArrayContext> _nestedContexts = new List<XlMarshalOperArrayContext>();
        bool _isExcel12v;    // Used for calls to Excel12 -- flags that returned native data should look different

        XlOper12* _pNative; // For managed -> native returns 
        // This points to the last OPER (and contained OPER array) that was marshaled
        // OPERs are re-allocated on every managed->native transition
        XlOper12* _pNativeStrings;
        XlOper12* _pNativeReferences;

        XlOper12* _pOperPointers; // Used for calls to Excel4v - points to the array of oper addresses

        public XlMarshalOperArrayContext(int rank, bool isExcel12v)
        {
            _rank = rank;
            _isExcel12v = isExcel12v;
        }

        // RESET

        // FREE
    }

    static class XlDirectMarshal
    {
        // Not cleaning up - we don't expect this to use a lot of memory, and think the pool of Excel calculation threads is stable
        readonly static ThreadLocal<XlMarshalContext> MarshalContext = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetMarshalContext() => MarshalContext.Value;

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

            //var types = methodInfo.Parameters.Select(pi => pi.DirectMarshalNativeType).ToList();
            //Type returnType = methodInfo.ReturnType?.DirectMarshalNativeType ?? typeof(void);
            //types.Add(returnType);
            ////return Expression.GetDelegateType(types.ToArray());
            //return DelegateCreator.MakeNewCustomDelegate(types.ToArray());
        }

        //// TODO: Cache
        //// TODO: Pre-defined delegates for IntPtr-only calls
        //static class DelegateCreator
        //{
        //    public static readonly Func<Type[], Type> MakeNewCustomDelegate = (Func<Type[], Type>)Delegate.CreateDelegate(
        //      typeof(Func<Type[], Type>),
        //      typeof(Expression).Assembly.GetType("System.Linq.Expressions.Compiler.DelegateHelpers").GetMethod(
        //        "MakeNewCustomDelegate",
        //        BindingFlags.NonPublic | BindingFlags.Static
        //      )
        //    );
        //    public static Type NewDelegateType(Type ret, params Type[] parameters)
        //    {
        //        var offset = parameters.Length;
        //        Array.Resize(ref parameters, offset + 1);
        //        parameters[offset] = ret;
        //        return MakeNewCustomDelegate(parameters);
        //    }
        //}

        // These are identifiers for xlfRegister types used in the pointer-only direct marshalling 
        public static readonly string XlTypeDoublePtr = "E";  // double*
        public static readonly string XlTypeXloper = "Q";

        public static readonly string XlTypeString = "D%";
    }


    //////////// TODO: Track registration string here?
    //////////class ConversionInfo
    //////////{
    //////////    public Type NativeType;
    //////////    public Type ManagedType;
    //////////    public Delegate ConvertN;    // Native -> Managed for parameters, and Managed -> Native for Return

    //////////    internal unsafe ConversionInfo(XlParameterInfo parameterInfo, bool isReturn, XlDirectMarshal directMarshal)
    //////////    {
    //////////        // TODO/DM Restructure the lookups here a bit better, or consolidate with SetTypeInfo12
    //////////        ManagedType = parameterInfo.DelegateParamType;

    //////////        switch (parameterInfo.XlType)
    //////////        {
    //////////            case "B":   // double
    //////////                NativeType = typeof(double);
    //////////                ConvertN = null;
    //////////                return;
    //////////            case "C%":   // string
    //////////                Debug.Assert(!isReturn);
    //////////                NativeType = typeof(IntPtr); // char*
    //////////                ConvertN = (Func<IntPtr, string>)(value => new string((char*)value));
    //////////                return;

    //////////            default:
    //////////                break;
    //////////        }

    //////////    }
    //////////}

    // These conversions for parameter and return values run with a MarshalContext for the thread in flight
    // Or we make open delegates and pass the context in explicitly
    // Or we use Expression.Lambda to glue the call in directly (i.e. Make these MethodCallExpressions)
    static unsafe class XlDirectConversions
    {
        public static MethodInfo DoubleToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleToXloperReturn));
        public static MethodInfo DoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrReturn));
        public static MethodInfo StringReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.StringReturn));
        public static MethodInfo ObjectToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectToXloperReturn));
        public static MethodInfo DateTimeToDoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeToDoublePtrReturn));

        // Param conversions are static - don't need context.
        public static MethodInfo DoublePtrParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrParam));
        public static MethodInfo StringParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.StringParam));

        public static MethodInfo DateTimeToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeToXloperReturn));
        public static MethodInfo DateTimeFromDoublePtrParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeFromDoublePtrParam));

        public static string Convert(char* value) => new string(value);


    }


    //
    //
    // NOTE: We also need to take care of the XlCall.Excel(...) call
    // It builds the context stack
}
