//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security;
using ExcelDna.Integration;

namespace ExcelDna.Loader
{
    internal class XlCallImpl
    {
        [DllImport("XLCALL32.DLL")]
        internal static extern int XLCallVer();

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string moduleName);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        [SuppressUnmanagedCodeSecurity]
        internal unsafe delegate int Excel12vDelegate(int xlfn, int count, XlOper12** ppOpers, XlOper12* pOperRes);
        internal static Excel12vDelegate Excel12v;

        /*
        ** Function number bits
        */
        public static readonly int xlCommand = 0x8000;
        public static readonly int xlSpecial = 0x4000;
        public static readonly int xlIntl = 0x2000;
        public static readonly int xlPrompt = 0x1000;

        /*
        ** Auxiliary function numbers
        **
        ** These functions are available only from the C API,
        ** not from the Excel macro language.
        */
        public static readonly int xlFree = (0 | xlSpecial);
        public static readonly int xlCoerce = (2 | xlSpecial);
        public static readonly int xlSheetId = (4 | xlSpecial);
        public static readonly int xlSheetNm = (5 | xlSpecial);
        public static readonly int xlGetHwnd = (8 | xlSpecial);
        public static readonly int xlGetName = (9 | xlSpecial);

        public static readonly int xlcAlert = (118 | xlCommand);
        public static readonly int xlcNew = (119 | xlCommand);
        public static readonly int xlcMessage = (122 | xlCommand);
        public static readonly int xlcEcho = (141 | xlCommand);
        public static readonly int xlcFileClose = (144 | xlCommand);
        public static readonly int xlcOnKey = (168 | xlCommand);
        public static readonly int xlcWorkbookInsert = (354 | xlCommand);

        public static readonly int xlfSetName = 88;
        public static readonly int xlfCaller = 89;
        public static readonly int xlfRegister = 149;
        public static readonly int xlfUnregister = 201;
        public static readonly int xlfRtd = 379;

        public static XlCall.XlReturn TryExcelImpl(int xlFunction, out object result, params object[] parameters)
        {
            if (Excel12v == null)
            {
                FetchExcel12EntryPt();
                if (Excel12v == null)
                {
                    result = null;
                    return XlCall.XlReturn.XlReturnFailed;
                }
            }

            return (XlCall.XlReturn)TryExcelImpl12(xlFunction, out result, parameters);
        }

        static void FetchExcel12EntryPt()
        {
            if (Excel12v == null)
            {
                try
                {
                    IntPtr hModuleProcess = GetModuleHandle(null);
                    IntPtr pfnExcel12v = GetProcAddress(hModuleProcess, "MdCallBack12");
                    if (pfnExcel12v != IntPtr.Zero)
                    {
                        Excel12v = Marshal.GetDelegateForFunctionPointer<Excel12vDelegate>(pfnExcel12v);
                    }
                }
                catch
                {
                }
            }
        }

        internal static void SetExcel12EntryPt(IntPtr pfnExcel12v)
        {
            Debug.Print("SetExcel12EntryPt called.");
            FetchExcel12EntryPt();
            if (Excel12v == null && pfnExcel12v != IntPtr.Zero)
            {
                Debug.Print("SetExcel12EntryPt - setting delegate.");
                Excel12v = Marshal.GetDelegateForFunctionPointer<Excel12vDelegate>(pfnExcel12v);
                Debug.Print("SetExcel12EntryPt - setting delegate OK? -  " + (Excel12v != null).ToString());
            }
        }

        private unsafe static int TryExcelImpl12(int xlFunction, out object result, params object[] parameters)
        {
            int xlReturn;

            // Set up the memory to hold the result from the call
            XlOper12 resultOper = new XlOper12();
            resultOper.xlType = XlType12.XlTypeEmpty;
            XlOper12* pResultOper = &resultOper;  // No need to pin for local struct

            // Special kind of ObjectArrayMarshaler for the parameters (rank 1)
            using (XlMarshalXlOperArrayContext paramMarshaler
                        = new XlMarshalXlOperArrayContext(1, true))
            {
                XlOper12** ppOperParameters = (XlOper12**)paramMarshaler.ObjectArrayReturn(parameters);
                xlReturn = Excel12v(xlFunction, parameters.Length, ppOperParameters, pResultOper);
            }

            // pResultOper now holds the result of the evaluated function
            // Get ObjectMarshaler for the return value
            result = XlMarshalContext.ObjectParam((IntPtr)pResultOper);

            // And free any memory allocated by Excel
            Excel12v(xlFree, 1, &pResultOper, (XlOper12*)IntPtr.Zero);

            return xlReturn;
        }

        public unsafe static IntPtr GetCurrentSheetId12()
        {
            // In a macro type function, xlSheetNm seems to return the Active sheet instead of the Current sheet.
            // So we first try to get the Current sheet from the caller.
            IntPtr retval = GetCallerSheetId12();
            if (retval != IntPtr.Zero)
                return retval;

            // Else we try the old way.
            XlOper12 SRef = new XlOper12();
            SRef.xlType = XlType12.XlTypeSReference;
            //SRef.srefValue.Count = 1;
            //SRef.srefValue.Reference.RowFirst = 1;
            //SRef.srefValue.Reference.RowLast = 1;
            //SRef.srefValue.Reference.ColumnFirst = 1;
            //SRef.srefValue.Reference.ColumnLast = 1;

            XlOper12 resultOper = new XlOper12();
            XlOper12* pResultOper = &resultOper;

            XlOper12* pSRef = &SRef;
            XlOper12** ppSRef = &(pSRef);
            int xlReturn;
            xlReturn = Excel12v(xlSheetNm, 1, ppSRef, pResultOper);
            if (xlReturn == 0)
            {
                XlOper12 resultRef = new XlOper12();
                XlOper12* pResultRef = &resultRef;
                xlReturn = Excel12v(xlSheetId, 1, (XlOper12**)&(pResultOper), pResultRef);

                // Done with pResultOper - Free
                Excel12v(xlFree, 1, &pResultOper, (XlOper12*)IntPtr.Zero);

                if (xlReturn == 0)
                {
                    if (resultRef.xlType == XlType12.XlTypeReference)
                    {
                        retval = resultRef.refValue.SheetId;
                    }
                    // Done with ResultRef - Free it too
                    Excel12v(xlFree, 1, &pResultRef, (XlOper12*)IntPtr.Zero);
                }
                // CONSIDER: As a small optimisation, we could combine the two calls the xlFree. But then we'd have to manage an array here.
            }
            return retval;
        }

        public unsafe static IntPtr GetCallerSheetId12()
        {
            IntPtr retval = IntPtr.Zero;
            XlOper12 resultOper = new XlOper12();
            XlOper12* pResultOper = &resultOper;
            int xlReturn;
            xlReturn = Excel12v(xlfCaller, 0, (XlOper12**)IntPtr.Zero, pResultOper);
            if (xlReturn == 0)
            {
                if (resultOper.xlType == XlType12.XlTypeReference)
                {
                    retval = resultOper.refValue.SheetId;
                    Excel12v(xlFree, 1, &pResultOper, (XlOper12*)IntPtr.Zero);
                }
            }
            return retval;
        }

    }
}
