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
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Globalization;

namespace ExcelDna.Integration
{
    [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil class.")]
    public class Excel
    {
        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.WindowHandle property.")]
        public static IntPtr WindowHandle
        {
            get { return ExcelDnaUtil.WindowHandle; }
        }

        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.Application property.")]
        public static object Application
        {
            get { return ExcelDnaUtil.Application; }
        }

        [Obsolete("Use ExcelDna.Integration.ExcelDnaUtil.IsInFunctionWizard property.")]
        public static bool IsInFunctionWizard()
        {
            return ExcelDnaUtil.IsInFunctionWizard();
        }
    }

	public class ExcelDnaUtil
	{
		private delegate bool EnumWindowsCallback(IntPtr hwnd, /*ref*/ IntPtr param);

		[DllImport("user32.dll")]
		private static extern int EnumWindows(EnumWindowsCallback callback, /*ref*/ IntPtr param);
		[DllImport("user32.dll")]
		private static extern IntPtr GetParent(IntPtr hwnd);
		[DllImport("user32.dll")]
		private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, /*ref*/ IntPtr param);
		[DllImport("user32.dll")]
		private static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowTextW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);
        [DllImport("Oleacc.dll")]
		private static extern int AccessibleObjectFromWindow(
			  IntPtr hwnd, uint dwObjectID, byte[] riid,
			  ref IntPtr ptr /*ppUnk*/);

		private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
		private static readonly byte[] IID_IDispatchBytes = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();

		internal static void Initialize()
		{
			// Need to get window in a macro context, else the call to get the Excel version fails.
            // Exception suppressor added here for HPC support and for RegSvr32 registration
            // - WindowHandle fails in these contexts.
            try
            {
                IntPtr unused = WindowHandle;
            }
            catch { }
		}

		private static IntPtr _hWndExcel = IntPtr.Zero;
		public static IntPtr WindowHandle
		{
			get
			{
				// CONSIDER: Process.GetCurrentProcess().MainWindowHandle;
				if (_hWndExcel == IntPtr.Zero)
				{
                    if (ExcelDnaUtil.ExcelVersion < 12)
                    {
                        // Only have the loword so far.
                        ushort loWord = (ushort)(double)XlCall.Excel(XlCall.xlGetHwnd);
                        EnumWindows(delegate(IntPtr hWndEnum, IntPtr param)
                            {
                                // Check the loWord
                                if (((uint)hWndEnum & 0x0000FFFF) == (uint)loWord)
                                {
                                    // Check the window class
                                    StringBuilder cname = new StringBuilder(256);
                                    GetClassNameW(hWndEnum, cname, cname.Capacity);
                                    if (cname.ToString() == "XLMAIN")
                                    {
                                        _hWndExcel = hWndEnum;
                                        return false;	// Stop enumerating
                                    }
                                }
                                return true;	// Continue enumerating
                            }, (IntPtr)0);
                    }
                    else
                    {
                        // TODO: 64-BIT - Check - This is clearly wrong.
                        _hWndExcel = (IntPtr)(int)(double)XlCall.Excel(XlCall.xlGetHwnd);
                    }
				}
				return _hWndExcel;
			}
		}

        [ThreadStatic]
        private static object _application;
        // CONSIDER: If we do load a COM Add-in for the Ribbon, should we keep that Application object 
        // around for convenience?
        // CONSIDER: Keep a WeakReference cache (per thread...), just for performance?
		public static object Application
		{
			get
			{
                // Check for a cached one set by a ComAddIn.
                if (_application != null) return _application;

                // Don't cache the one we get from the Window, it keeps Excel alive!
                object application;
                application = GetApplicationFromWindow();
                if (application == null)
                {
                    // I assume it failed because there was no workbook open
                    // Now make workbook with VBA sheet, according to some Google post

                    // CONSIDER: Alternative of sending WM_USER+18 to Excel - KB 147573
                    //           And trying to retrieve Excel from the ROT using GetActiveObject
                    //           Concern then is whether it is the right instance of the Excel.Application for this process.


                    // DOCUMENT: Under some circumstances, the C API and Automation interfaces are not available.
                    //  This happens when there is no Workbook open in Excel.
                    // We try a (possible) test for whether we can call the C API.
                    object output;
                    XlCall.XlReturn result = XlCall.TryExcel(XlCall.xlGetName, out output);
                    if (result == XlCall.XlReturn.XlReturnFailed)
                    {
                        // no plan for getting Application.
                        throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object.");
                    }

                    // Create new workbook with the right stuff
                    XlCall.Excel(XlCall.xlcEcho, false);
                    XlCall.Excel(XlCall.xlcNew, 5);
                    XlCall.Excel(XlCall.xlcWorkbookInsert, 6);

                    application = GetApplicationFromWindow();

                    // Clean up
                    XlCall.Excel(XlCall.xlcFileClose, false);
                    XlCall.Excel(XlCall.xlcEcho, true);
                }
                return application;
			}
            internal set
            {
                // Should only be set from Com Add-in connect.
                _application = value;
            }
		}

		private static object GetApplicationFromWindow()
		{
			// This is Andrew Whitechapel's plan for getting the Application object.
			// It does not work when there are no Workbooks open.
			IntPtr hWndMain = WindowHandle;
			IntPtr hWndChild = IntPtr.Zero;
			EnumChildWindows(hWndMain, delegate(IntPtr hWndEnum, IntPtr param)
			{
				// Check the window class
				StringBuilder cname = new StringBuilder(256);
				GetClassNameW(hWndEnum, cname, cname.Capacity);
				if (cname.ToString() == "EXCEL7")
				{
					hWndChild = hWndEnum;
					return false;	// Stop enumerating
				}
				return true;	// Continue enumerating
			}, (IntPtr)0);
            if (hWndChild != (IntPtr)0)
            {
                IntPtr pUnk = (IntPtr)0;
                int hr = AccessibleObjectFromWindow(
                        hWndChild, OBJID_NATIVEOM,
                        IID_IDispatchBytes, ref pUnk);
                if (hr >= 0)
                {
                    object obj = Marshal.GetObjectForIUnknown(pUnk);
                    Marshal.Release(pUnk);

                    object app = obj.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, obj, null, new CultureInfo(1033));
                    Marshal.ReleaseComObject(obj);

                    //							object ver = app.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, app, null);
                    return app;
                }
            }
			return null;
		}

		public static bool IsInFunctionWizard()
		{
			// TODO: Handle the Find and Replace dialog
			//       for international versions.
			IntPtr hWndMain = WindowHandle;
			bool inFunctionWizard = false;
            StringBuilder cname = new StringBuilder(256);
            StringBuilder title = new StringBuilder(256);
			EnumWindows(delegate(IntPtr hWndEnum, IntPtr param)
			{
				// Check the window class
				GetClassNameW(hWndEnum, cname, cname.Capacity);
				if (cname.ToString().StartsWith("bosa_sdm_XL"))
				{
					if (GetParent(hWndEnum) == hWndMain)
					{
						GetWindowTextW(hWndEnum, title, title.Capacity);
						if (!title.ToString().Contains("Replace"))
							inFunctionWizard = true; // will also work for older versions where paste box had no title
						return false;	// Stop enumerating
					}
				}
				return true;	// Continue enumerating
			}, (IntPtr)0);
			return inFunctionWizard;
		}

		//private static double _xlVersion = 0;
		//public static double ExcelVersion
		//{
		//    get
		//    {
		//        if (_xlVersion == 0)
		//        {
		//            object versionString;
		//            versionString = XlCall.Excel(XlCall.xlfGetWorkspace, 2);
		//            double version;
		//            bool parseOK = double.TryParse((string)versionString, out version);
		//            if (!parseOK)
		//            {
		//                // Might be locale problem 
		//                // and Excel 12 returns versionString with "." as decimal sep.
		//                //  ->  microsoft.public.excel.sdk thread
		//                //      Excel4(xlfGetWorkspace, &version, 1, & arg) - Excel2007 Options 
		//                //      Dec 12, 2006
		//                parseOK = double.TryParse((string)versionString,
		//                            System.Globalization.NumberStyles.AllowDecimalPoint,
		//                            System.Globalization.NumberFormatInfo.InvariantInfo,
		//                            out version);
		//            }
		//            if (!parseOK)
		//            {
		//                version = 0.99;
		//            }
		//            _xlVersion = version;
		//        }
		//        return _xlVersion;
		//    }
		//}

		// Updated for International Excel - Thanks to Martin Drecher
		private static double _xlVersion = 0;
		public static double ExcelVersion
		{
			get
			{
				if (_xlVersion == 0)
				{
                    object versionObject;
                    XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlfGetWorkspace, out versionObject, 2);
                    if (retval == XlCall.XlReturn.XlReturnSuccess)
                    {
                        string versionString = System.Text.RegularExpressions.Regex.Match((string)versionObject, "^\\d+(\\.\\d+)?").Value;
                        double version;
                        if (!String.IsNullOrEmpty(versionString) &&
                            Double.TryParse(versionString, System.Globalization.NumberStyles.Any,
                                            System.Globalization.CultureInfo.InvariantCulture, out version))
                        {
                            _xlVersion = version;
                        }
                        else
                        {
                            _xlVersion = 0.99;
                        }
                    }
                    else
                    {
                        // Maybe running on a Cluster
                        object isRunningOnClusterResult;
                        retval = XlCall.TryExcel(XlCall.xlRunningOnCluster, out isRunningOnClusterResult);
                        if (retval == XlCall.XlReturn.XlReturnSuccess && (isRunningOnClusterResult is string))
                        {
                            // TODO: How to get the real version here...?
                            _xlVersion = 14.0;
                        }
                    }
				}
				return _xlVersion;
			}
		} 

        private static ExcelLimits _xlLimits;
        public static ExcelLimits ExcelLimits
        {
            get
            {
                if (_xlLimits == null)
                {
                    _xlLimits = new ExcelLimits();
                    if (ExcelVersion < 12.0)
                    {
                        _xlLimits.MaxRows = 65536;
                        _xlLimits.MaxColumns = 256;
                        _xlLimits.MaxArguments = 30;
                        _xlLimits.MaxStringLength = 255;
                    }
                    else
                    {
                        _xlLimits.MaxRows = 1048576;
                        _xlLimits.MaxColumns = 16384;
                        _xlLimits.MaxArguments = 255;
                        _xlLimits.MaxStringLength = 32767;
                    }
                }
                return _xlLimits;
            }
        }
	}

    public class ExcelLimits
    {
		private int _maxRows;
		public int MaxRows
		{
			get { return _maxRows; }
			internal set { _maxRows = value; }
		}

		private int _maxColumns;
		public int MaxColumns
		{
			get { return _maxColumns; }
			internal set { _maxColumns = value; }
		}

		private int _maxArguments;
		public int MaxArguments
		{
			get { return _maxArguments; }
			internal set { _maxArguments = value; }
		}

		private int _maxStringLength;
		public int MaxStringLength
		{
			get { return _maxStringLength; }
			internal set { _maxStringLength = value; }
		}
	}

}
