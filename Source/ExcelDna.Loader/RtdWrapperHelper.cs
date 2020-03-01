using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelDna.Loader
{
    unsafe class RtdWrapperHelper
    {
        public static MethodInfo GetRtdWrapperMethod()
        {
            return typeof(RtdWrapperHelper).GetMethod("RtdWrapper");
        }

        readonly string _progId;
        readonly object _rtdWrapperOptions;
        readonly XlOper12* _progIdXloper12;
        readonly XlOper12* _emptyStringXloper12;
        readonly XlOper12* _errorValueXloper12;
        
        // TODO: Need a .NET 2.0 equivalent of ThreadLocal to make this thread-safe
        readonly XlOper12* _resultXloper12;  // Not Thread-Safe !?
        readonly XlOper12** _callParams;

        public RtdWrapperHelper(string progId, object rtdWrapperOptions)
        {
            _progId = progId;
            _rtdWrapperOptions = rtdWrapperOptions;

            // TODO: Can't easily use the regular marshallers, because we want to opt out of the memory management
            // TODO: But we should refactor this conversion code a bit

            int size = Marshal.SizeOf(typeof(XlOper12));
            int stringSize = Marshal.SizeOf(typeof(XlString12)) + ((XlString12.MaxLength - 1) /* 1 char is in Data[1] */ * 2 /* 2 bytes per char */);
            
            _progIdXloper12 = (XlOper12*)Marshal.AllocCoTaskMem(size);
            _progIdXloper12->xlType = XlType12.XlTypeString;
            _progIdXloper12->pstrValue = (XlString12*)Marshal.AllocCoTaskMem(stringSize);
            XlString12 * pdest = _progIdXloper12->pstrValue;
            ushort charCount = (ushort)Math.Min(progId.Length, XlString12.MaxLength);
            fixed (char* psrc = progId)
            {
                char* ps = psrc;
                char* pd = pdest->Data;
                for (int k = 0; k < charCount; k++)
                {
                    *(pd++) = *(ps++);
                }
            }
            pdest->Length = charCount;

            _emptyStringXloper12 = (XlOper12*)Marshal.AllocCoTaskMem(size);
            _emptyStringXloper12->xlType = XlType12.XlTypeString;
            _emptyStringXloper12->pstrValue = (XlString12*)Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(XlString12)));
            _emptyStringXloper12->pstrValue->Length = 0;

            _errorValueXloper12 = (XlOper12*)Marshal.AllocCoTaskMem(size);
            _errorValueXloper12->errValue = 15; // ExcelErrorValue
            _errorValueXloper12->xlType = XlType12.XlTypeError;

            _resultXloper12 = (XlOper12*)Marshal.AllocCoTaskMem(size);
            _resultXloper12->xlType = XlType12.XlTypeEmpty;

            _callParams = (XlOper12**)Marshal.AllocCoTaskMem(IntPtr.Size * 6);
            _callParams[0] = _progIdXloper12;
            _callParams[1] = _emptyStringXloper12;
        }

        // This is the function we register with Excel
        // All the IntPtrs are XLOPER12*
        // TODO: Test static vs instance delegates ?
        public unsafe XlOper12* RtdWrapper(XlOper12* topic1, XlOper12* topic2, XlOper12* topic3, XlOper12* topic4)
        {
            _callParams[2] = topic1;
            _callParams[3] = topic2;
            _callParams[4] = topic3;
            _callParams[5] = topic4;

            int xlReturn = XlCallImpl.Excel12v(XlCallImpl.xlfRtd, 6, _callParams, _resultXloper12);
            if (xlReturn == 0) // xlReturnSuccess)
            {
                _resultXloper12->xlType |= XlType12.XlBitXLFree;
                return _resultXloper12;
            }
            else
            {
                return _errorValueXloper12;
            }
        }
    }
}
