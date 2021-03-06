//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelDna.Loader
{
    // We have one XlMarshalContext per thread
    // It is never collected
    unsafe class XlMarshalContext
    {
        // TODO/DM: Test again that these are actually meaningful - it seems strange
        readonly static object boxedZero = 0.0;
        readonly static object boxedOne = 1.0;

        readonly static object excelEmpty = ExcelEmpty.Value;

        // These are fixed size, and could be allocated as a single struct or block.
        // Strings of any length, in Xloper or direct, using max length fixed buffer
        readonly XlString12* _pStringBufferReturn;
        readonly double* _pDoubleReturn; // Also used for DateTime
        readonly short* _pBoolReturn;

        // All the in-place Xloper types
        readonly XlOper12* _pXloperReturn;

        // // Used for single-element array return, allowing allocation-free return in this case
        // readonly XlOper12* _pXloperArraySingletonReturn;

        readonly XlMarshalDoubleArrayContext _rank1DoubleArrayContext;
        readonly XlMarshalDoubleArrayContext _rank2DoubleArrayContext;

        readonly XlMarshalXlOperArrayContext _rank1OperArrayContext;
        readonly XlMarshalXlOperArrayContext _rank2OperArrayContext;

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

            _rank1OperArrayContext = new XlMarshalXlOperArrayContext(1, false);
            _rank2OperArrayContext = new XlMarshalXlOperArrayContext(2, false);
        }

        internal void FreeMemory()
        {
            _rank1OperArrayContext.Reset(true);
            _rank2OperArrayContext.Reset(true);
        }

        // RULE: Return conversions must not throw exceptions (they might run in the exception handler)
        // RULE: Param conversions can throw exceptions

        public IntPtr ObjectReturn(object ManagedObj)
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
            if (ManagedObj is double d)
            {
                _pXloperReturn->numValue = d;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is string str)
            {
                _pXloperReturn->pstrValue = (XlString12*)StringReturn(str);
                _pXloperReturn->xlType = XlType12.XlTypeString;
            }
            else if (ManagedObj is DateTime dt)
            {
                return DateTimeToXloperReturn(dt);
            }
            else if (ManagedObj is bool b)
            {
                return BoolToXloperReturn(b);
            }
            else if (ManagedObj is object[] arr1)
            {
                // Redirect to the ObjectArray Marshaler
                // CONSIDER: This might cause some memory to get stuck, 
                // since the memory for the array marshaler is not the same as for this
                return ObjectArray1Return(arr1);
            }
            else if (ManagedObj is object[,] arr2)
            {
                // Redirect to the ObjectArray Marshaler
                // CONSIDER: This might cause some memory to get stuck, 
                // since the memory for the array marshaler is not the same as for this
                return ObjectArray2Return(arr2);
            }
            else if (ManagedObj is double[] d1)
            {
                return DoubleArray1Return(d1);
            }
            else if (ManagedObj is double[,] d2)
            {
                return DoubleArray2Return(d2);
            }
            else if (ManagedObj is ExcelError err)
            {
                _pXloperReturn->errValue = (short)err;
                _pXloperReturn->xlType = XlType12.XlTypeError;
            }
            else if (ManagedObj is ExcelMissing)
            {
                _pXloperReturn->xlType = XlType12.XlTypeMissing;
            }
            else if (ManagedObj is ExcelEmpty)
            {
                _pXloperReturn->xlType = XlType12.XlTypeEmpty;
            }
            else if (ManagedObj is short s)
            {
                _pXloperReturn->numValue = s;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is Missing)
            {
                _pXloperReturn->xlType = XlType12.XlTypeMissing;
            }
            else if (ManagedObj is int i)
            {
                _pXloperReturn->numValue = i;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is uint ui)
            {
                _pXloperReturn->numValue = ui;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is byte byt)
            {
                _pXloperReturn->numValue = byt;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is ushort us)
            {
                _pXloperReturn->numValue = us;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is decimal dec)
            {
                _pXloperReturn->numValue = (double)dec;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is float f)
            {
                _pXloperReturn->numValue = f;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is long l)
            {
                _pXloperReturn->numValue = l;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is ulong ul)
            {
                _pXloperReturn->numValue = ul;
                _pXloperReturn->xlType = XlType12.XlTypeNumber;
            }
            else if (ManagedObj is ExcelAsyncHandleNative ah)
            {
                // This code is not actually used, since the ExcelAsyncHandle is only passed into
                // XlCall.Excel param array, so marshaled by the object array marshaler.

                _pXloperReturn->bigData.hData = ah.Handle;
                _pXloperReturn->bigData.cbData = IntPtr.Size;
                _pXloperReturn->xlType = XlType12.XlTypeBigData;
            }
            // CONSIDER: Reimplement in this class (needs extra memory management)?
            else if (ManagedObj is ExcelReference xlref)
            {
                // To avoid extra memory management in this class, wrap in an array and let the array marshaler deal with the reference.
                // TODO/DM: It would be better to have the extra copy of the code here, or abstract the reference memory context too and share
                object[] refArray = new object[1];
                refArray[0] = xlref;
                XlOper12* pArray = (XlOper12*)ObjectArray1Return(refArray);

                // Pick reference out of the returned array.
                return (IntPtr)pArray->arrayValue.pOpers;
            }
            else
            {
                // Default error return
                _pXloperReturn->errValue = (int)ExcelError.ExcelErrorValue;
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

        public IntPtr DoubleArray1Return(double[] doubles)
        {
            return _rank1DoubleArrayContext.DoubleArrayReturn(doubles);
        }

        public IntPtr DoubleArray2Return(double[,] doubles)
        {
            return _rank2DoubleArrayContext.DoubleArrayReturn(doubles);
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

        public IntPtr Int16ToXloperReturn(short i)
        {
            _pXloperReturn->numValue = i;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr UInt16ToXloperReturn(ushort i)
        {
            _pXloperReturn->numValue = i;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr Int32ToXloperReturn(int i)
        {
            _pXloperReturn->numValue = i;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr Int64ToXloperReturn(long i)
        {
            _pXloperReturn->numValue = i;
            _pXloperReturn->xlType = XlType12.XlTypeNumber;
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr DecimalToXloperReturn(decimal d)
        {
            _pXloperReturn->numValue = (double)d;
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
                _pXloperReturn->errValue = (int)ExcelError.ExcelErrorValue;
                _pXloperReturn->xlType = XlType12.XlTypeError;
            }
            return (IntPtr)_pXloperReturn;
        }

        public IntPtr BoolPtrReturn(bool b)
        {
            *_pBoolReturn = b ? (short)1 : (short)0;
            return (IntPtr)_pBoolReturn;
        }

        // Input parameter conversions (also for XlCall.Excel return values) - static, no context
        public static object ObjectParam(IntPtr pNativeData)
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
                    managed = (ExcelError)pOper->errValue;
                    break;
                case XlType12.XlTypeMissing:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = System.Reflection.Missing.Value;
                    managed = ExcelMissing.Value;
                    break;
                case XlType12.XlTypeEmpty:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = null;
                    managed = ExcelEmpty.Value;
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
                                array[i, j] = ObjectParam((IntPtr)oper);
                            }
                        }
                    }
                    managed = array;
                    break;
                case XlType12.XlTypeReference:
                    object /*ExcelReference*/ r;
                    if (pOper->refValue.pMultiRef == (XlOper12.XlMultiRef12*)IntPtr.Zero)
                    {
                        r = new ExcelReference(0, 0, 0, 0, pOper->refValue.SheetId);
                    }
                    else
                    {
                        ushort numAreas = *(ushort*)pOper->refValue.pMultiRef;
                        // XlOper12.XlRectangle12* pAreas = (XlOper12.XlRectangle12*)((uint)pOper->refValue.pMultiRef + 4 /* FieldOffset for XlRectangles */);
                        XlOper12.XlRectangle12* pAreas = (XlOper12.XlRectangle12*)((byte*)(pOper->refValue.pMultiRef) + 4 /* FieldOffset for XlRectangles */);
                        if (numAreas == 1)
                        {

                            r = new ExcelReference(
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
                            r = new ExcelReference(areas, pOper->refValue.SheetId);
                        }
                    }
                    managed = r;
                    break;
                case XlType12.XlTypeSReference:
                    IntPtr sheetId = XlCallImpl.GetCurrentSheetId12();
                    object /*ExcelReference*/ sref;
                    sref = new ExcelReference(
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
            var managed = ObjectParam(pNativeData);
            return (object[])XlMarshalXlOperArrayContext.ObjectArrayParam(managed, 1);
        }

        public static object[,] ObjectArray2Param(IntPtr pNativeData)
        {
            var managed = ObjectParam(pNativeData);
            return (object[,])XlMarshalXlOperArrayContext.ObjectArrayParam(managed, 2);
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

        public static bool BoolPtrParam(IntPtr pb)
        {
            return *(short*)pb == 1;
        }

        public static short DoublePtrToInt16Param(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the Int16 range
            //          This code will raise an exception and we have to handle it or return immediately or something
            return checked((short)(Math.Round(*(double*)pNativeData, MidpointRounding.ToEven)));
        }

        public static ushort DoublePtrToUInt16Param(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the Int16 range
            //          This code will raise an exception and we have to handle it or return immediately or something
            return checked((ushort)(Math.Round(*(double*)pNativeData, MidpointRounding.ToEven)));
        }

        public static int DoublePtrToInt32Param(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the Int32 range
            //          This code will raise an exception and we have to handle it or return immediately or something
            return checked((int)(Math.Round(*(double*)pNativeData, MidpointRounding.ToEven)));
        }

        public static long DoublePtrToInt64Param(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the Int64 range
            //          This code will raise an exception and we have to handle it or return immediately or something
            return checked((long)(Math.Round(*(double*)pNativeData, MidpointRounding.ToEven)));
        }
        public static decimal DoublePtrToDecimalParam(IntPtr pNativeData)
        {
            // TODO/DM: Check what happens when we're outside the Decimal range
            //          This code will raise an exception and we have to handle it or return immediately or something
            return (decimal)*((double*)pNativeData);
        }

        public static ExcelAsyncHandle AsyncHandleParam(IntPtr pNativeData)
        {
            // Make a nice object from the native OPER
            XlOper12* pOper = (XlOper12*)pNativeData;
            // Ignore any Free flags
            XlType12 type = pOper->xlType & ~XlType12.XlBitXLFree & ~XlType12.XlBitDLLFree;
            if (type == XlType12.XlTypeBigData)
            {
                XlOper12.XlBigData bigData = pOper->bigData;
                return new ExcelAsyncHandleNative(bigData.hData);
            }

            throw new ArgumentException("Unexpected XlOper type for AsyncHandle: " + type, nameof(pNativeData));
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
                throw new InvalidOperationException("Damaged XlMarshalDoubleArray rank");
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

    unsafe class XlMarshalXlOperArrayContext : IDisposable
    {
        int _rank;
        // These used for array return
        List<XlMarshalXlOperArrayContext> _nestedInstances = new List<XlMarshalXlOperArrayContext>();
        bool _isExcel12v;    // Used for calls to Excel12 -- flags that returned native data should look different

        // All of these are XlOper12*
        IntPtr _pNative; // For managed -> native returns 
        // This points to the last OPER (and contained OPER array) that was marshaled
        // OPERs are re-allocated on every managed->native transition
        IntPtr _pNativeStrings;
        IntPtr _pNativeReferences;

        IntPtr _pOperPointers; // Used for calls to Excel4v - points to the array of oper addresses

        public XlMarshalXlOperArrayContext(int rank, bool isExcel12v)
        {
            _rank = rank;
            _isExcel12v = isExcel12v;
        }

        // IDisposable implementation.
        // Instances that are created explicitly (not via marshaling during function baking)
        // are the only ones explicitly disposed. 
        // These isntances created in calls from XlCallImpl.
        private bool disposed = false;
        public void Dispose()
        {
            // Debug.Print("Disposing XlObjectArray12Marshaler with id {0} for thread {1}", id, System.Threading.Thread.CurrentThread.ManagedThreadId);
            Dispose(true);
            GC.SuppressFinalize(this);  // TODO/DM We really are allocating native memory, but does it matter whether we finalize? It might we never lose track of the object and can clean up nicely
        }

        // Also called to clean up the instance on every return call...
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                Reset(disposing);
            }
            disposed = true;
        }

        ~XlMarshalXlOperArrayContext()
        {
            Dispose(false);
        }

        // Called for disposal and for reset on every call to ManagedToNative.
        public void Reset(bool disposeNested)
        {
            if (disposeNested)
            {
                // Clean up the nested Instances
                foreach (XlMarshalXlOperArrayContext m in _nestedInstances)
                {
                    m.Reset(true);
                }
                _nestedInstances.Clear();
            }

            if (_pNative != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_pNative);
                _pNative = IntPtr.Zero;
            }

            if (_pNativeStrings != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_pNativeStrings);
                _pNativeStrings = IntPtr.Zero;
            }

            if (_pNativeReferences != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_pNativeReferences);
                _pNativeReferences = IntPtr.Zero;
            }

            if (_pOperPointers != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_pOperPointers);
                _pOperPointers = IntPtr.Zero;
            }
        }

        public IntPtr ObjectArrayReturn(object ManagedObj)
        {
            Reset(true);

            // DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!
            if (ManagedObj == null)
                return IntPtr.Zero;

            // CONSIDER: Managing memory differently
            // Here we allocate and clear when the next array is returned
            // we might also return XLOPER and have xlFree called back.

            // TODO: Remove duplication - due to fixed / pointer interaction

            int rows;
            int columns; // those in the returned array
            int rowBase;
            int columnBase;
            if (_rank == 1)
            {
                object[] objects = (object[])ManagedObj;

                rows = 1;
                rowBase = 0;
                columns = objects.Length;
                columnBase = objects.GetLowerBound(0);
            }
            else if (_rank == 2)
            {
                object[,] objects = (object[,])ManagedObj;

                rows = objects.GetLength(0);
                rowBase = objects.GetLowerBound(0);
                columns = objects.GetLength(1);
                columnBase = objects.GetLowerBound(0);
            }
            else
            {
                throw new InvalidOperationException("Damaged XlMarshalXlOperArrayContext rank");
            }

            int cbNativeStrings = 0;
            int numReferenceOpers = 0;
            int numReferences = 0;

            // Allocate native space
            int cbNative = Marshal.SizeOf(typeof(XlOper12)) +               // OPER that is returned
                           Marshal.SizeOf(typeof(XlOper12)) * (rows * columns);    // Array of OPER inside the result
            _pNative = Marshal.AllocCoTaskMem(cbNative);

            // Set up returned OPER
            XlOper12* pOper = (XlOper12*)_pNative;
            // Excel chokes badly on empty arrays (e.g. crash in function wizard) - rather return the default erro value, #VALUE!
            if (rows * columns == 0)
            {
                pOper->errValue = (ushort)ExcelError.ExcelErrorValue;
                pOper->xlType = XlType12.XlTypeError;
            }
            else
            {

                pOper->xlType = XlType12.XlTypeArray;
                pOper->arrayValue.Rows = rows;
                pOper->arrayValue.Columns = columns;
                pOper->arrayValue.pOpers = ((XlOper12*)_pNative + 1);
            }
            // This loop won't be entered in the empty-array case (rows * columns == 0)
            for (int i = 0; i < rows * columns; i++)
            {
                // Get the right object out of the array
                object obj;
                if (_rank == 1)
                {
                    obj = ((object[])ManagedObj)[columnBase + i];
                }
                else
                {
                    int row = i / columns;
                    int column = i % columns;
                    obj = ((object[,])ManagedObj)[rowBase + row, columnBase + column];
                }

                // Get the right pOper
                pOper = (XlOper12*)_pNative + i + 1;

                // Set up the oper from the object
                if (obj is double d)
                {
                    pOper->numValue = d;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is string str)
                {
                    // We count all of the string lengths, 
                    cbNativeStrings += (Marshal.SizeOf(typeof(XlString12)) + ((Math.Min(str.Length, XlString12.MaxLength) - 1) /* 1 char already in XlString */) * 2 /* 2 bytes per char */);
                    // mark the Oper as a string, and
                    // later allocate memory and return to fix pointers
                    pOper->xlType = XlType12.XlTypeString;
                }
                else if (obj is DateTime dt)
                {
                    try
                    {
                        pOper->numValue = dt.ToOADate();
                        pOper->xlType = XlType12.XlTypeNumber;
                    }
                    catch
                    {
                        // This is a case where we have a date that cannot be converted to an OleAutomation date, e.g. year < 0100
                        // Certainly it is not a valid date in Excel (no dates before 1900 are valid)
                        // But we must not crash - return #VALUE instead
                        pOper->errValue = (int)ExcelError.ExcelErrorValue;
                        pOper->xlType = XlType12.XlTypeError;
                    }
                }
                else if (obj is ExcelError err)
                {
                    pOper->errValue = (short)err;
                    pOper->xlType = XlType12.XlTypeError;
                }
                else if (obj is ExcelMissing)
                {
                    pOper->xlType = XlType12.XlTypeMissing;
                }
                else if (obj is ExcelEmpty)
                {
                    pOper->xlType = XlType12.XlTypeEmpty;
                }
                else if (obj is ExcelAsyncHandleNative ah)
                {
                    pOper->bigData.hData = ah.Handle;
                    pOper->bigData.cbData = IntPtr.Size;
                    pOper->xlType = XlType12.XlTypeBigData;
                }
                else if (obj is bool b)
                {
                    pOper->boolValue = b ? 1 : 0;
                    pOper->xlType = XlType12.XlTypeBoolean;
                }
                else if (obj is byte byt)
                {
                    pOper->numValue = byt;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is sbyte sbyt)
                {
                    pOper->numValue = sbyt;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is short sh)
                {
                    pOper->numValue = sh;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is ushort ush)
                {
                    pOper->numValue = ush;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is int ii)
                {
                    pOper->numValue = ii;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is uint ui)
                {
                    pOper->numValue = ui;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is long l)
                {
                    pOper->numValue = l;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is ulong ul)
                {
                    pOper->numValue = ul;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is decimal dec)
                {
                    pOper->numValue = (double)dec;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is float fl)
                {
                    pOper->numValue = fl;
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is ExcelReference xlref)
                {
                    pOper->xlType = XlType12.XlTypeReference;
                    // First we count all of these, 
                    // later allocate memory and return to fix pointers
                    numReferenceOpers++;
                    numReferences += xlref.GetRectangleCount();
                }
                else if (obj is object[] arr1)
                {
                    var nested = new XlMarshalXlOperArrayContext(1, false);
                    _nestedInstances.Add(nested);
                    XlOper12* pNested = (XlOper12*)nested.ObjectArrayReturn(arr1);
                    if (pNested->xlType == XlType12.XlTypeArray)
                    {
                        pOper->xlType = XlType12.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType12.XlTypeError;
                        pOper->errValue = (int)ExcelError.ExcelErrorValue;
                    }
                }
                else if (obj is object[,] arr2)
                {
                    var nested = new XlMarshalXlOperArrayContext(2, false);
                    _nestedInstances.Add(nested);
                    XlOper12* pNested = (XlOper12*)nested.ObjectArrayReturn(arr2);
                    if (pNested->xlType == XlType12.XlTypeArray)
                    {
                        pOper->xlType = XlType12.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0,0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType12.XlTypeError;
                        pOper->errValue = (int)ExcelError.ExcelErrorValue;
                    }
                }
                else if (obj is double[] doubles)
                {
                    object[] objects = new object[doubles.Length];
                    Array.Copy(doubles, objects, doubles.Length);

                    var nested = new XlMarshalXlOperArrayContext(1, false);
                    _nestedInstances.Add(nested);
                    XlOper12* pNested = (XlOper12*)nested.ObjectArrayReturn(objects);
                    if (pNested->xlType == XlType12.XlTypeArray)
                    {
                        pOper->xlType = XlType12.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType12.XlTypeError;
                        pOper->errValue = (int)ExcelError.ExcelErrorValue;
                    }
                }
                else if (obj is double[,] doubles2)
                {
                    object[,] objects = new object[doubles2.GetLength(0), doubles2.GetLength(1)];
                    Array.Copy(doubles2, objects, doubles2.GetLength(0) * doubles2.GetLength(1));

                    var nested = new XlMarshalXlOperArrayContext(2, false);
                    _nestedInstances.Add(nested);
                    XlOper12* pNested = (XlOper12*)nested.ObjectArrayReturn(objects);
                    if (pNested->xlType == XlType12.XlTypeArray)
                    {
                        pOper->xlType = XlType12.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0,0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType12.XlTypeError;
                        pOper->errValue = (int)ExcelError.ExcelErrorValue;
                    }
                }
                else if (obj is Missing)
                {
                    pOper->xlType = XlType12.XlTypeMissing;
                }
                else if (obj == null)
                {
                    // DOCUMENT: I return Empty for nulls inside the Array, 
                    // which is not consistent with what happens in other settings.
                    // In particular not consistent with the results of the XlObjectMarshaler
                    // (which is not called when a null is returned,
                    // and interpreted as ExcelErrorNum in Excel)
                    // This works well for xlSet though.
                    pOper->xlType = XlType12.XlTypeEmpty;
                }
                else
                {
                    // Default error return
                    pOper->xlType = XlType12.XlTypeError;
                    pOper->errValue = (int)ExcelError.ExcelErrorValue;
                }
            } // end of first pass

            // Now handle strings
            if (cbNativeStrings > 0)
            {
                // Allocate room for all the strings
                _pNativeStrings = Marshal.AllocCoTaskMem(cbNativeStrings);
                // Go through the Opers and set each string
                char* pCurrent = (char*)_pNativeStrings;
                for (int i = 0; i < rows * columns; i++)
                {
                    // Get the corresponding oper
                    pOper = (XlOper12*)_pNative + i + 1;
                    if (pOper->xlType == XlType12.XlTypeString)
                    {
                        // Get the string from the managed array
                        string str;
                        if (_rank == 1)
                        {
                            str = (string)((object[])ManagedObj)[i];
                        }
                        else
                        {
                            int row = i / columns;
                            int column = i % columns;
                            str = (string)((object[,])ManagedObj)[rowBase + row, columnBase + column];
                        }

                        XlString12* pdest = (XlString12*)pCurrent;
                        pOper->pstrValue = pdest;
                        ushort charCount = (ushort)Math.Min(str.Length, XlString12.MaxLength);
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
                        // Increment pointer within allocated memory
                        pCurrent += charCount + 1;
                    }
                }
            }

            // Now handle references
            if (numReferenceOpers > 0)
            {
                // Allocate room for all the references
                int cbNativeReferences = numReferenceOpers * 4 /* sizeof ushort + packing to get to field offset */
                                         + numReferences * Marshal.SizeOf(typeof(XlOper12.XlRectangle12));
                _pNativeReferences = Marshal.AllocCoTaskMem(cbNativeReferences);
                IntPtr pCurrent = _pNativeReferences;
                // Go through the Opers and set each reference
                int refOperIndex = 0;
                for (int i = 0; i < rows * columns && refOperIndex < numReferenceOpers; i++)
                {
                    // Get the corresponding oper
                    pOper = (XlOper12*)_pNative + i + 1;
                    if (pOper->xlType == XlType12.XlTypeReference)
                    {
                        // Get the reference from the managed array
                        ExcelReference r;
                        if (_rank == 1)
                        {
                            r = (ExcelReference)((object[])ManagedObj)[i];
                        }
                        else
                        {
                            int row = i / columns;
                            int column = i % columns;
                            r = (ExcelReference)((object[,])ManagedObj)[rowBase + row, columnBase + column];
                        }

                        int refCount = r.GetRectangleCount();
                        int numBytes = 4 /* sizeof ushort + packing to get to field offset */
                                       + refCount * Marshal.SizeOf(typeof(XlOper12.XlRectangle12));


                        IntPtr sheetId = r.SheetId;
                        int[][] rects = r.GetRectangles();
                        int rectCount = rects.GetLength(0);

                        pOper->xlType = XlType12.XlTypeReference;
                        pOper->refValue.SheetId = sheetId;

                        pOper->refValue.pMultiRef = (XlOper12.XlMultiRef12*)pCurrent;
                        pOper->refValue.pMultiRef->Count = (ushort)rectCount;

                        XlOper12.XlRectangle12* pRectangles = &pOper->refValue.pMultiRef->Rectangles;

                        for (int ir = 0; ir < rectCount; ir++)
                        {
                            pRectangles[ir].RowFirst = rects[ir][0];
                            pRectangles[ir].RowLast = rects[ir][1];
                            pRectangles[ir].ColumnFirst = rects[ir][2];
                            pRectangles[ir].ColumnLast = rects[ir][3];
                        }

                        // Unchecked keyword here is redundant (it's the default for C#), 
                        // but makes clear that we rely on the overflow.
                        // Also - numBytes must be int and not long, else we get numeric promotion and a mess again!
                        pCurrent = IntPtr.Size == 4 ?
                            new IntPtr(unchecked(pCurrent.ToInt32() + (int)numBytes)) :
                            new IntPtr(unchecked(pCurrent.ToInt64() + numBytes));
                        refOperIndex++;
                    }
                }
            }

            if (!_isExcel12v)
            {
                // For big allocations, ensure that Excel allows us to free the memory
                if (rows * columns * 16 + cbNativeStrings + numReferences * 16 > 65535)
                    pOper->xlType |= XlType12.XlBitDLLFree;

                // We are done
                return _pNative;
            }
            else
            {
                // For the Excel12v call, we need to return an array
                // which will contain the pointers to the Opers.
                int cbOperPointers = columns * Marshal.SizeOf(typeof(XlOper12*));
                _pOperPointers = Marshal.AllocCoTaskMem(cbOperPointers);
                XlOper12** pOpers = (XlOper12**)_pOperPointers;
                for (int i = 0; i < columns; i++)
                {
                    pOpers[i] = (XlOper12*)_pNative + i + 1;
                }
                return _pOperPointers;
            }
        }

        public static object ObjectArrayParam(object managed, int rank)
        {
            // Duplication here, because the types are different and wrapped in fixed blocks
            if (rank == 1)
            {
                if (managed == null || !(managed is object[,]))
                {
                    return new object[1] { managed };
                }
                else // managed is object[,]: turn first row (or column) into object[]
                {
                    object[] array;
                    object[,] all = (object[,])managed;
                    int rows = all.GetLength(0);
                    int columns = all.GetLength(1);

                    if (columns == 1)
                    {
                        // Take the one and only column as the array
                        array = new object[rows];
                        for (int i = 0; i < rows; i++)
                        {
                            array[i] = all[i, 0];
                        }
                    }
                    else
                    {
                        // Take first row only
                        array = new object[columns];
                        for (int j = 0; j < columns; j++)
                        {
                            array[j] = all[0, j];
                        }
                    }
                    return array;
                }
            }
            else if (rank == 2)
            {
                if (managed == null || !(managed is object[,]))
                {
                    return new object[,] { { managed } };
                }
                else // managed is object[,]
                {
                    return managed;
                }
            }
            else
            {
                throw new InvalidOperationException("Damaged XlMarshalXlOperArrayContext rank");
            }
        }
    }

    static class XlMarshalConversions
    {
        // These conversions for return values run with a MarshalContext for the thread in flight
        public static MethodInfo ObjectReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectReturn));
        public static MethodInfo ObjectArray1Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray1Return));
        public static MethodInfo ObjectArray2Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray2Return));
        public static MethodInfo DoubleArray1Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray1Return));
        public static MethodInfo DoubleArray2Return = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray2Return));
        public static MethodInfo DoubleToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleToXloperReturn));
        public static MethodInfo DoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrReturn));
        public static MethodInfo StringReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.StringReturn));
        public static MethodInfo DateTimeToDoublePtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeToDoublePtrReturn));
        public static MethodInfo DateTimeToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeToXloperReturn));
        public static MethodInfo BoolPtrReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.BoolPtrReturn));
        public static MethodInfo BoolToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.BoolToXloperReturn));
        public static MethodInfo Int16ToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.Int16ToXloperReturn));
        public static MethodInfo UInt16ToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.UInt16ToXloperReturn));
        public static MethodInfo Int32ToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.Int32ToXloperReturn));
        public static MethodInfo Int64ToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.Int64ToXloperReturn));
        public static MethodInfo DecimalToXloperReturn = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DecimalToXloperReturn));

        // Param conversions are static - don't need context.
        public static MethodInfo ObjectParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectParam));
        public static MethodInfo ObjectArray1Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray1Param));
        public static MethodInfo ObjectArray2Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.ObjectArray2Param));
        public static MethodInfo DoubleArray1Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray1Param));
        public static MethodInfo DoubleArray2Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoubleArray2Param));
        public static MethodInfo DoublePtrParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrParam));
        public static MethodInfo StringParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.StringParam));
        public static MethodInfo DateTimeFromDoublePtrParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DateTimeFromDoublePtrParam));
        public static MethodInfo BoolPtrParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.BoolPtrParam));
        public static MethodInfo DoublePtrToInt16Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrToInt16Param));
        public static MethodInfo DoublePtrToUInt16Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrToUInt16Param));
        public static MethodInfo DoublePtrToInt32Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrToInt32Param));
        public static MethodInfo DoublePtrToInt64Param = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrToInt64Param));
        public static MethodInfo DoublePtrToDecimalParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.DoublePtrToDecimalParam));
        public static MethodInfo AsyncHandleParam = typeof(XlMarshalContext).GetMethod(nameof(XlMarshalContext.AsyncHandleParam));
    }
}
