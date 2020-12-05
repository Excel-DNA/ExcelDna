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

        readonly static object excelEmpty = IntegrationMarshalHelpers.GetExcelEmptyValue();  // NOTE: What is the timing for this static initialization?

        // These are fixed size, and could be allocated as a single struct or block.
        // Strings of any length, in Xloper or direct, using max length fixed buffer
        readonly XlString12* _pStringBufferReturn;
        readonly double* _pDoubleReturn; // Also used for DateTime
        readonly short* _pBoolReturn;

        // All the in-place Xloper types
        readonly XlOper12* _pXloperReturn;

        // Used for single-element array return, allowing allocation-free return in this case
        readonly XlOper12* _pXloperArraySingletonReturn;

        readonly XlMarshalDoubleArrayContext _rank1DoubleArrayContext;
        readonly XlMarshalDoubleArrayContext _rank2DoubleArrayContext;

        readonly XlMarshalOperArrayContext _rank1OperArrayContext;
        readonly XlMarshalOperArrayContext _rank2OperArrayContext;

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

        public IntPtr ObjectToXloperReturn(object ManagedObj)
        {
            // We maintain compatible behaviour with the CustomMarshalling, which would return null pointers directly (without calling marshalling)
            // TODO/DM: DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!
            if (ManagedObj == null)
                return IntPtr.Zero;

            // CONSIDER: Managing memory differently
            // Here we allocate and clear when the next object is returned
            // we might also return XLOPER with the right bits set and have xlFree called back (which we do for large object arrays)

            // Debug.Print("XlObject12Marshaler {0} - Marshaling for thread {1} ", instanceId, System.Threading.Thread.CurrentThread.ManagedThreadId);

            // CONSIDER: Use TypeHandle of Type.GetTypeCode(type) lookup instead of if/else?
            if (ManagedObj is double)
            {
                _pXloperReturn->numValue = (double)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is string)
            {
                _pXloperReturn->pstrValue = (XlString12*)StringReturn((string)ManagedObj);
                _pXloperReturn->xlType = XlType12.XlTypeString;
            }
            else if (ManagedObj is DateTime)
            {
                return DateTimeToXloperReturn((DateTime)ManagedObj);
            }
            else if (ManagedObj is bool)
            {
                return BoolToXloperReturn((bool)ManagedObj);
            }
            else if (ManagedObj is object[])
            {
                // Redirect to the ObjectArray Marshaler
                // CONSIDER: This might cause some memory to get stuck, 
                // since the memory for the array marshaler is not the same as for this
                return ObjectArray1Return((object[])ManagedObj);
            }
            else if (ManagedObj is object[,])
            {
                // Redirect to the ObjectArray Marshaler
                // CONSIDER: This might cause some memory to get stuck, 
                // since the memory for the array marshaler is not the same as for this
                return ObjectArray2Return((object[,])ManagedObj);
            }
            else if (ManagedObj is double[])
            {
                return DoubleArray1Return((double[])ManagedObj);
            }
            else if (ManagedObj is double[,])
            {
                return DoubleArray2Return((double[,])ManagedObj);
            }
            else if (IntegrationMarshalHelpers.IsExcelErrorObject(ManagedObj))
            {
                _pXloperReturn->errValue = IntegrationMarshalHelpers.ExcelErrorGetValue(ManagedObj);
                _pXloperReturn->xlType = XlType12.XlTypeError;
            }
            else if (IntegrationMarshalHelpers.IsExcelMissingObject(ManagedObj))
            {
                _pXloperReturn->xlType = XlType12.XlTypeMissing;
            }
            else if (IntegrationMarshalHelpers.IsExcelEmptyObject(ManagedObj))
            {
                _pXloperReturn->xlType = XlType12.XlTypeEmpty;
            }
            else if (ManagedObj is short)
            {
                _pXloperReturn->numValue = (short)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is System.Reflection.Missing)
            {
                _pXloperReturn->xlType = XlType12.XlTypeMissing;
            }
            else if (ManagedObj is int)
            {
                _pXloperReturn->numValue = (int)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is uint)
            {
                _pXloperReturn->numValue = (uint)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is byte)
            {
                _pXloperReturn->numValue = (byte)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is ushort)
            {
                _pXloperReturn->numValue = (ushort)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is decimal)
            {
                _pXloperReturn->numValue = (double)((decimal)ManagedObj);
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is float)
            {
                _pXloperReturn->numValue = (float)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is long)
            {
                _pXloperReturn->numValue = (long)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is ulong)
            {
                _pXloperReturn->numValue = (ulong)ManagedObj;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (IntegrationMarshalHelpers.IsExcelAsyncHandleNativeObject(ManagedObj))
            {
                // This code is not actually used, since the ExcelAsyncHandle is only passed into
                // XlCall.Excel param array, so marshaled byt the object array marshaler.
                IntPtr handle = IntegrationMarshalHelpers.GetExcelAsyncHandleNativeHandle(ManagedObj);

                _pXloperReturn->bigData.hData = handle;
                _pXloperReturn->bigData.cbData = IntPtr.Size;
                _pXloperReturn->xlType = XlType12.XlTypeBigData;
            }
            // CONSIDER: Reimplement in this class (needs extra memory management)?
            else if (IntegrationMarshalHelpers.IsExcelReferenceObject(ManagedObj))
            {
                // To avoid extra memory management in this class, wrap in an array and let the array marshaler deal with the reference.
                // TODO/DM: It would be better to have the extra copy of the code here, or abstract the reference memory context too and share
                object[] refArray = new object[1];
                refArray[0] = ManagedObj;
                XlOper12* pArray = (XlOper12*)ObjectArray1Return(refArray);

                // Pick reference out of the returned array.
                return (IntPtr)pArray->arrayValue.pOpers;
            }
            else
            {
                // Default error return
                _pXloperReturn->errValue = IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                _pXloperReturn->xlType = XlType12.XlTypeError;
            }
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr ObjectArray1Return(object[] objects)
        {
            return _rank1OperArrayContext.ObjectArrayReturn(objects);
        }

        public IntPtr ObjectArray2Return(object[,] objects)
        {
            return _rank2OperArrayContext.ObjectArrayReturn(objects);
        }

        public IntPtr DoubleToXloperReturn(double d)
        {
            _pXloperReturn->numValue = d;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr BoolToXloperReturn(bool b)
        {
            _pXloperReturn->numValue = b ? 1 : 0;
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
            // We maintain compatible behaviour with the CustomMarshaling, which would return null pointers directly (without calling marshalling)
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

        public IntPtr DoubleArray1Return(double[] doubles)
        {
            return _rank1DoubleArrayContext.DoubleArrayReturn(doubles);
        }

        public IntPtr DoubleArray2Return(double[,] doubles)
        {
            return _rank2DoubleArrayContext.DoubleArrayReturn(doubles);
        }

        // Input parameter conversions (also for XlCall.Excel return values) - static, no context
        public static object XloperToObjectParam(IntPtr pNativeData)
        {
            // Make a nice object from the native OPER
            object managed;
            XlOper12* pOper = (XlOper12*)pNativeData;
            // Ignore any Free flags
            XlType12 type = pOper->xlType & ~XlType12.XlBitXLFree & ~XlType12.XlBitDLLFree;
            switch (type)
            {
                case XlType12.XlTypeNumber:
                    double val = pOper->numValue;
                    if (val == 0.0)
                        managed = boxedZero;
                    else if (val == 1.0)
                        managed = boxedOne;
                    else
                        managed = val;
                    break;
                case XlType12.XlTypeString:
                    XlString12* pString = pOper->pstrValue;
                    managed = new string(pString->Data, 0, pString->Length);
                    break;
                case XlType12.XlTypeBoolean:
                    managed = pOper->boolValue == 1;
                    break;
                case XlType12.XlTypeError:
                    managed = IntegrationMarshalHelpers.GetExcelErrorObject(pOper->errValue);
                    break;
                case XlType12.XlTypeMissing:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = System.Reflection.Missing.Value;
                    managed = IntegrationMarshalHelpers.GetExcelMissingValue();
                    break;
                case XlType12.XlTypeEmpty:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = null;
                    managed = excelEmpty; // IntegrationMarshalHelpers.GetExcelEmptyValue();
                    break;
                case XlType12.XlTypeArray:
                    int rows = pOper->arrayValue.Rows;
                    int columns = pOper->arrayValue.Columns;
                    object[,] array = new object[rows, columns];
                    // TODO: Initialize as ExcelEmpty?
                    XlOper12* opers = (XlOper12*)pOper->arrayValue.pOpers;
                    for (int i = 0; i < rows; i++)
                    {
                        int rowStart = i * columns;
                        for (int j = 0; j < columns; j++)
                        {
                            int pos = rowStart + j;
                            XlOper12* oper = opers + pos;
                            // Fast-path for some cases
                            if (oper->xlType == XlType12.XlTypeEmpty)
                            {
                                array[i, j] = excelEmpty;
                            }
                            else if (oper->xlType == XlType12.XlTypeNumber)
                            {
                                double dval = oper->numValue;
                                if (dval == 0.0)
                                    array[i, j] = boxedZero;
                                else if (dval == 1.0)
                                    array[i, j] = boxedOne;
                                else
                                    array[i, j] = dval;
                            }
                            else
                            {
                                array[i, j] = XloperToObjectParam((IntPtr)oper);
                            }
                        }
                    }
                    managed = array;
                    break;
                case XlType12.XlTypeReference:
                    object /*ExcelReference*/ r;
                    if (pOper->refValue.pMultiRef == (XlOper12.XlMultiRef12*)IntPtr.Zero)
                    {
                        r = IntegrationMarshalHelpers.CreateExcelReference(0, 0, 0, 0, pOper->refValue.SheetId);
                    }
                    else
                    {
                        ushort numAreas = *(ushort*)pOper->refValue.pMultiRef;
                        // XlOper12.XlRectangle12* pAreas = (XlOper12.XlRectangle12*)((uint)pOper->refValue.pMultiRef + 4 /* FieldOffset for XlRectangles */);
                        XlOper12.XlRectangle12* pAreas = (XlOper12.XlRectangle12*)((byte*)(pOper->refValue.pMultiRef) + 4 /* FieldOffset for XlRectangles */);
                        if (numAreas == 1)
                        {

                            r = IntegrationMarshalHelpers.CreateExcelReference(
                                pAreas[0].RowFirst, pAreas[0].RowLast,
                                pAreas[0].ColumnFirst, pAreas[0].ColumnLast, pOper->refValue.SheetId);
                        }
                        else
                        {
                            int[][] areas = new int[numAreas][];
                            for (int i = 0; i < numAreas; i++)
                            {
                                XlOper12.XlRectangle12 rect = pAreas[i];
                                int[] area = new int[4] { rect.RowFirst, rect.RowLast,
                                                          rect.ColumnFirst, rect.ColumnLast };
                                areas[i] = area;
                            }
                            r = IntegrationMarshalHelpers.CreateExcelReference(areas, pOper->refValue.SheetId);
                        }
                    }
                    managed = r;
                    break;
                case XlType12.XlTypeSReference:
                    IntPtr sheetId = XlCallImpl.GetCurrentSheetId12();
                    object /*ExcelReference*/ sref;
                    sref = IntegrationMarshalHelpers.CreateExcelReference(
                                            pOper->srefValue.Reference.RowFirst,
                                            pOper->srefValue.Reference.RowLast,
                                            pOper->srefValue.Reference.ColumnFirst,
                                            pOper->srefValue.Reference.ColumnLast,
                                            sheetId /*Current sheet (not active sheet)*/);
                    managed = sref;
                    break;
                case XlType12.XlTypeInt: // Never passed from Excel to a UDF! int32 in XlOper12
                    managed = (double)pOper->intValue;
                    break;
                default:
                    // unheard of !!
                    managed = null;
                    break;
            }
            return managed;
        }

        public static object[] ObjectArray1Param(IntPtr pNativeData)
        {
            return (object[])XlMarshalOperArrayContext.ObjectArrayParam(pNativeData, 1);
        }

        public static object[,] ObjectArray2Param(IntPtr pNativeData)
        {
            return (object[,])XlMarshalOperArrayContext.ObjectArrayParam(pNativeData, 2);
        }

        public static double[] DoubleArray1Param(IntPtr pDoubles)
        {
            return (double[])XlMarshalDoubleArrayContext.DoubleArrayParam(pDoubles, 1);
        }

        public static double[,] DoubleArray2Param(IntPtr pDoubles)
        {
            return (double[,])XlMarshalDoubleArrayContext.DoubleArrayParam(pDoubles, 2);
        }

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
        IntPtr _pNative; // For managed -> native returns

        public XlMarshalDoubleArrayContext(int rank)
        {
            _rank = rank;
        }

        unsafe public IntPtr DoubleArrayReturn(object doubleArray)
        {
            // CONSIDER: Checking checking object type
            // CONSIDER: Managing memory differently
            // Here we allocate and clear when the next array is returned
            // we might also return XLOPER and have xlFree called back.

            // If array is too big!?, we just truncate

            // TODO: Remove duplication - due to fixed / pointer interaction

            if (doubleArray == null)
                return IntPtr.Zero; // #NUM!

            Marshal.FreeCoTaskMem(_pNative);
            _pNative = IntPtr.Zero;

            int rows;
            int columns;
            if (_rank == 1)
            {
                double[] doubles = (double[])doubleArray;

                rows = 1;
                columns = doubles.Length;

                // Guard against invalid arrays - with no columns.
                // Just return null, which Excel will turn into #NUM
                if (columns == 0)
                    return IntPtr.Zero;

                fixed (double* src = doubles)
                {
                    AllocateFP12AndCopy(src, rows, columns);
                }
            }
            else if (_rank == 2)
            {
                double[,] doubles = (double[,])doubleArray;

                rows = doubles.GetLength(0);
                columns = doubles.GetLength(1);

                // Guard against invalid arrays - with no rows or no columns.
                // Just return null, which Excel will turn into #NUM
                if (rows == 0 || columns == 0)
                    return IntPtr.Zero;

                fixed (double* src = doubles)
                {
                    AllocateFP12AndCopy(src, rows, columns);
                }
            }
            else
            {
                throw new InvalidOperationException("Damaged XlDoubleArrayMarshaler rank");
            }

            // CONSIDER: If large, mark and deal with xlDllFree

            return _pNative;
        }

        unsafe private void AllocateFP12AndCopy(double* pSrc, int rows, int columns)
        {
            // CONSIDER: Fast memcpy: http://stackoverflow.com/questions/1715224/very-fast-memcpy-for-image-processing
            // CONSIDER: https://connect.microsoft.com/VisualStudio/feedback/details/766977/il-bytecode-method-cpblk-badly-implemented-by-x86-clr
            XlFP12* pFP;

            int size = Marshal.SizeOf(typeof(XlFP12)) +
                Marshal.SizeOf(typeof(double)) * (rows * columns - 1); // room for one double is already in FP12 struct
            _pNative = Marshal.AllocCoTaskMem(size);

            pFP = (XlFP12*)_pNative;
            pFP->Rows = rows;
            pFP->Columns = columns;
            int count = rows * columns;
            // Fast copy
            CopyDoubles(pSrc, pFP->Values, count);
        }


        public static object DoubleArrayParam(IntPtr pNativeData, int rank)
        {
            object result;
            XlFP12* pFP = (XlFP12*)pNativeData;

            // Duplication here, because the types are different and wrapped in fixed blocks
            if (rank == 1)
            {
                double[] array;
                if (pFP->Columns == 1)
                {
                    // Take the one and only column as the array
                    array = new double[pFP->Rows];
                }
                else
                {
                    // Take only the first row of the array.
                    array = new double[pFP->Columns];
                }
                // Copy works for either case, due to in-memory layout!
                fixed (double* dest = array)
                {
                    CopyDoubles(pFP->Values, dest, array.Length);
                }
                result = array;
            }
            else if (rank == 2)
            {
                double[,] array = new double[pFP->Rows, pFP->Columns];
                fixed (double* dest = array)
                {
                    CopyDoubles(pFP->Values, dest, array.Length);
                }
                result = array;
            }
            else
            {
                Debug.Fail("Damaged XlDoubleArray12Marshaler rank");
                throw new InvalidOperationException("Damaged XlDoubleArray12Marshaler rank");
            }
            return result;
        }

        static void CopyDoubles(double* pSrc, double* pDest, int count)
        {
            for (int i = 0; i < count; i++)
            {
                pDest[i] = pSrc[i];
            }
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

        internal IntPtr ObjectArrayReturn(object managedObj)
        {
            throw new NotImplementedException();
        }

        public static object ObjectArrayParam(IntPtr pNativeData, int rank)
        {
            throw new NotImplementedException();
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
        public static readonly string XlTypeXloper = "Q";
        public static readonly string XlTypeXloperAllowRef = "U";
        public static readonly string XlTypeDoublePtr = "E";        // double*
        public static readonly string XlTypeString = "D%";          // XLSTRING12
        public static readonly string XlTypeDoubleArray = "K%";     // FP12*
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
        public static MethodInfo ObjectToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectToXloperReturn));
        public static MethodInfo ObjectArray1Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray1Return));
        public static MethodInfo ObjectArray2Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray2Return));
        public static MethodInfo DoubleArray1Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray1Return));
        public static MethodInfo DoubleArray2Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray2Return));
        public static MethodInfo DoubleToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleToXloperReturn));
        public static MethodInfo DoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrReturn));
        public static MethodInfo StringReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.StringReturn));
        public static MethodInfo DateTimeToDoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeToDoublePtrReturn));

        // Param conversions are static - don't need context.
        public static MethodInfo XloperToObjectParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.XloperToObjectParam));
        public static MethodInfo ObjectArray1Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray1Param));
        public static MethodInfo ObjectArray2Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray2Param));
        public static MethodInfo DoubleArray1Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray1Param));
        public static MethodInfo DoubleArray2Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray2Param));
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
