/*
  Copyright (C) 2005-2011 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Loader
{
	internal class XlCallImpl
	{
        [DllImport("XLCALL32.DLL")]
        internal static extern int XLCallVer();

		[DllImport("XLCALL32.DLL")]
		private static extern unsafe int Excel4v(int xlfn, XlOper* pOperRes, int count, XlOper** ppOpers);

		[DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string moduleName);

		[DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

		[UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private unsafe delegate int Excel12vDelegate(int xlfn, int count, XlOper12** ppOpers, XlOper12* pOperRes);
		private static Excel12vDelegate Excel12v;

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
        public static readonly int xlcWorkbookInsert = (354 | xlCommand);

        public static readonly int xlfSetName = 88;
        public static readonly int xlfRegister = 149;
        public static readonly int xlfUnregister = 201;

        public unsafe static int TryExcelImpl(int xlFunction, out object result, params object[] parameters)
        {
            if (XlAddIn.XlCallVersion < 12)
            {
                return TryExcelImpl4(xlFunction, out result, parameters);
            }

            // Else Excel 12+
            if (Excel12v == null )
            {
                FetchExcel12EntryPt();
                if (Excel12v == null)
                {
                    result = null;
                    return 32; /*XlCall.XlReturn.XlReturnFailed*/
                }
            }

            return TryExcelImpl12(xlFunction, out result, parameters);
        }

        private static void FetchExcel12EntryPt()
        {
            if (Excel12v == null)
            {
                try
                {
                    IntPtr hModuleProcess = GetModuleHandle(null);
                    IntPtr pfnExcel12v = GetProcAddress(hModuleProcess, "MdCallBack12");
                    if (pfnExcel12v != IntPtr.Zero)
                    {
                        Excel12v = (Excel12vDelegate)Marshal.GetDelegateForFunctionPointer(pfnExcel12v, typeof(Excel12vDelegate));
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
            if (Excel12v == null && pfnExcel12v != null)
            {
                Debug.Print("SetExcel12EntryPt - setting delegate.");
                Excel12v = (Excel12vDelegate)Marshal.GetDelegateForFunctionPointer(pfnExcel12v, typeof(Excel12vDelegate));
                Debug.Print("SetExcel12EntryPt - setting delegate OK? -  " + (Excel12v != null).ToString());
            }
        }

        private unsafe static int TryExcelImpl4(int xlFunction, out object result, params object[] parameters)
		{
            int xlReturn;

            // Set up the memory to hold the result from the call
            XlOper resultOper = new XlOper();
            resultOper.xlType = XlType.XlTypeEmpty;
            XlOper* pResultOper = &resultOper;  // No need to pin for local struct

            // Special kind of ObjectArrayMarshaler for the parameters (rank 1)
            using (XlObjectArrayMarshaler paramMarshaler = new XlObjectArrayMarshaler(1, true))
            {
                XlOper** ppOperParameters = (XlOper**)paramMarshaler.MarshalManagedToNative(parameters);
                xlReturn = Excel4v(xlFunction, pResultOper, parameters.Length, ppOperParameters);
            }

            // pResultOper now holds the result of the evaluated function
            // Get ObjectMarshaler for the return value
            ICustomMarshaler m = XlObjectMarshaler.GetInstance("");
            result = m.MarshalNativeToManaged((IntPtr)pResultOper);
            // And free any memory allocated by Excel
            Excel4v(xlFree, (XlOper*)IntPtr.Zero, 1, &pResultOper);
        
            return xlReturn;
        }

        private unsafe static int TryExcelImpl12(int xlFunction, out object result, params object[] parameters)
        {
            int xlReturn;

            // Set up the memory to hold the result from the call
            XlOper12 resultOper = new XlOper12();
            resultOper.xlType = XlType12.XlTypeEmpty;
            XlOper12* pResultOper = &resultOper;  // No need to pin for local struct

            // Special kind of ObjectArrayMarshaler for the parameters (rank 1)
            using (XlObjectArray12Marshaler paramMarshaler = new XlObjectArray12Marshaler(1, true))
            {
                XlOper12** ppOperParameters = (XlOper12**)paramMarshaler.MarshalManagedToNative(parameters);
                xlReturn = Excel12v(xlFunction, parameters.Length, ppOperParameters, pResultOper);
            }

            // pResultOper now holds the result of the evaluated function
            // Get ObjectMarshaler for the return value
            ICustomMarshaler m = XlObject12Marshaler.GetInstance("");
            result = m.MarshalNativeToManaged((IntPtr)pResultOper);
            // And free any memory allocated by Excel
            Excel12v(xlFree, 1, &pResultOper, (XlOper12*)IntPtr.Zero);

            return xlReturn;
        }

        public unsafe static IntPtr GetCurrentSheetId4()
        {
            IntPtr retval = IntPtr.Zero;
            XlOper SRef = new XlOper();
            SRef.xlType = XlType.XlTypeSReference;
            //SRef.srefValue.Count = 1;
            //SRef.srefValue.Reference.RowFirst = 1;
            //SRef.srefValue.Reference.RowLast = 1;
            //SRef.srefValue.Reference.ColumnFirst = 1;
            //SRef.srefValue.Reference.ColumnLast = 1;

            XlOper resultOper = new XlOper();
            XlOper* pResultOper = &resultOper;

            XlOper* pSRef = &SRef;
            XlOper** ppSRef = &(pSRef);
            int xlReturn;
            xlReturn = Excel4v(xlSheetNm, pResultOper, 1, ppSRef);
            if (xlReturn == 0)
            {
                XlOper resultRef = new XlOper();
                XlOper* pResultRef = &resultRef;
                xlReturn = Excel4v(xlSheetId, pResultRef, 1, (XlOper**)&(pResultOper));

                // Done with pResultOper - Free
                Excel4v(xlFree, (XlOper*)IntPtr.Zero, 1, &pResultOper);

                if (xlReturn == 0)
                {
                    if (resultRef.xlType == XlType.XlTypeReference)
                    {
                        return resultRef.refValue.SheetId;
                    }
                    // Done with ResultRef - Free it too
                    Excel4v(xlFree, (XlOper*)IntPtr.Zero, 1, &pResultRef);
                }
            }
            return retval;
        }

        public unsafe static IntPtr GetCurrentSheetId12()
        {
            IntPtr retval = IntPtr.Zero;
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
	}
}
