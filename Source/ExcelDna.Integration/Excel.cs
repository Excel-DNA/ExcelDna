/*
  Copyright (C) 2005-2012 Govert van Drimmelen

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
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Globalization;

namespace ExcelDna.Integration
{
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
                // Return cached value if we have one
				if (_hWndExcel != IntPtr.Zero) return _hWndExcel;

                //// Try to get it the easy way from the Process info
                //// (doesn't work if Excel is not visible yet)
                //_hWndExcel = Process.GetCurrentProcess().MainWindowHandle;
                //if (_hWndExcel != IntPtr.Zero) return _hWndExcel;

                // Else get via the C API
                ushort loWord;
                if (ExcelDnaUtil.ExcelVersion >= 12)
                {
                    // Excel 2007+ - should get whole handle ...
                    IntPtr hWnd = (IntPtr) (int) (double) XlCall.Excel(XlCall.xlGetHwnd);
                    // ... but this call is not reliable - so check if we got an Excel window ...
                    if (IsAnExcelWindow(hWnd))
                    {
                        _hWndExcel = hWnd;
                        return _hWndExcel;
                    }

                    // Else go the lo-Word check
                    loWord = (ushort)hWnd;
                }
                else
                {
                    // Excel < 2007 - only have the loword so far.
                    loWord = (ushort)(double) XlCall.Excel(XlCall.xlGetHwnd);
                }
			    _hWndExcel = FindAnExcelWindow(loWord);

                // Might still be null...!
                if (_hWndExcel == IntPtr.Zero)
                {
                    Debug.Print("Failed to get Excel WindowHandle.");
                }
			    return _hWndExcel;
			}
		}

        // Check if hWnd refers to a Window of class "XLMAIN" indicating and Excel top-level window.
        static bool IsAnExcelWindow(IntPtr hWnd)
        {
            StringBuilder cname = new StringBuilder(256);
            GetClassNameW(hWnd, cname, cname.Capacity);
            return cname.ToString() == "XLMAIN";
        }

        // Try to find an Excel window with window handle that matches the passed lo word.
        static IntPtr FindAnExcelWindow(ushort hWndLoWord)
        {
            IntPtr hWnd = IntPtr.Zero;
            EnumWindows(delegate(IntPtr hWndEnum, IntPtr param)
            {
                // Check the loWord
                if (((uint)hWndEnum & 0x0000FFFF) == (uint)hWndLoWord &&
                    IsAnExcelWindow(hWndEnum))
                {
                    hWnd = hWndEnum;
                    return false;  // Stop enumerating
                }
                return true;  // Continue enumerating
            }, (IntPtr)0);
            return hWnd;
        }

	    [ThreadStatic]
        private static object _application;
		public static object Application
		{
			get
			{
                // Check for a cached one set by a ComAddIn.
                if (_application != null) return _application;


                // Get main window as well as we can.
                IntPtr hWndMain = WindowHandle;
                if (hWndMain == IntPtr.Zero) return null;   // This is a problematic error case!

                // Don't cache the one we get from the Window, it keeps Excel alive!
                object application;
                application = GetApplicationFromWindow(hWndMain);
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
                        // no plan for getting Application (we're probably on a different thread?)
                        throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object.");
                    }

                    // Create new workbook with the right stuff
                    XlCall.Excel(XlCall.xlcEcho, false);
                    XlCall.Excel(XlCall.xlcNew, 5);
                    XlCall.Excel(XlCall.xlcWorkbookInsert, 6);

                    // Try again
                    application = GetApplicationFromWindow(hWndMain);

                    // Clean up
                    XlCall.Excel(XlCall.xlcFileClose, false);
                    XlCall.Excel(XlCall.xlcEcho, true);
                }
                return application;
			}
            internal set
            {
                // Should only be set from Com Add-in connect, and only on the main thread.
                _application = value;
            }
		}

		private static object GetApplicationFromWindow(IntPtr hWndMain)
		{
			// This is Andrew Whitechapel's plan for getting the Application object.
			// It does not work when there are no Workbooks open.
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
			}, IntPtr.Zero);
            if (hWndChild != IntPtr.Zero)
            {
                IntPtr pUnk = IntPtr.Zero;
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
                            Double.TryParse(versionString, NumberStyles.Any, CultureInfo.InvariantCulture, out version))
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
                        else
                        {
                            // Maybe we are loading a COM Server or an RTD Server before the add-in is loaded.
                            try
                            {
                                object xlApp = ExcelDnaUtil.Application;
                                object result = xlApp.GetType().InvokeMember("Version",
                                                                             System.Reflection.BindingFlags.GetProperty,
                                                                             null, xlApp, null, new CultureInfo(1033));
                                _xlVersion = double.Parse((string)result, NumberStyles.Any, CultureInfo.InvariantCulture);
                            }
                            catch (Exception ex)
                            {
                                Debug.Print("Failed to get Excel version - " + ex);
                            }
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
