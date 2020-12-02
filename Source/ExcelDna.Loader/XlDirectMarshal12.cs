using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelDna.Loader
{
    // We have one XlMarshalContext per thread
    public unsafe class XlMarshalContext
    {
        // Strings of any length, in Xloper or direct, using max length fixed buffer
        XlString12* _pStringBufferReturn;
        double* _pDateTimeReturn;

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
            _pDateTimeReturn = (double*)Marshal.AllocCoTaskMem(size);

            // XloperReturn
            size = Marshal.SizeOf(typeof(XlOper12));
            _pXloperReturn = (XlOper12*)Marshal.AllocCoTaskMem(size);

            _rank1DoubleArrayContext = new XlMarshalDoubleArrayContext(1);
            _rank2DoubleArrayContext = new XlMarshalDoubleArrayContext(2);

            _rank1OperArrayContext = new XlMarshalOperArrayContext(1, false);
            _rank2OperArrayContext = new XlMarshalOperArrayContext(2, false);
        }
    }

    public unsafe class XlMarshalDoubleArrayContext
    {
        int _rank;
        XlFP12* _pNative; // For managed -> native returns

        public XlMarshalDoubleArrayContext(int rank)
        {
            _rank = rank;
        }
    }

    public unsafe class XlMarshalOperArrayContext
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

    public class XlDirectMarshal
    {
        // TODO/DM: private ThreadLocal<object> myFoo = new ThreadLocal<object>(() => new object());

        [ThreadStatic]
        static XlMarshalContext _marshalContext;

        public static object MarshalContext
        {
            get
            {
                if (_marshalContext == null)
                    _marshalContext = new XlMarshalContext();
                return _marshalContext;
            }
        }

        readonly static object boxedZero = 0.0;
        readonly static object boxedOne = 1.0;
        //        readonly object excelEmpty = IntegrationMarshalHelpers.GetExcelEmptyValue();

        // Given a delegate and information about the intended export, we instantiate and return the exportable delegate
        // The delegate captures this (singleton) object and then reads the MarshalContext from the ThreadLocal
        static Delegate CreateExport()
        {
            throw new NotImplementedException();
        }

    }

    // These conversions for parameter and return values run with a MarshalContext for the thread in flight
    public static class XlDirectConversions
    {

    }
}
