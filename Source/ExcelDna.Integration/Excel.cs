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
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using ExcelDna.Logging;

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
			// Exception suppressor added here for HPC support and for RegSvr32 registration
            // - WindowHandle fails in these contexts.
            try
            {
                IntPtr unused = WindowHandle;
                _mainThreadId = Thread.CurrentThread.ManagedThreadId;
            }
            catch (Exception ex)
            { 
                Debug.WriteLine("Error during ExcelDnaUtil.Initialize: " + ex);
                // Just suppress otherwise
            }
		}

		private static IntPtr _hWndExcel = IntPtr.Zero;
		public static IntPtr WindowHandle
		{
			get
			{
                // Return cached value if we have one
				if (_hWndExcel != IntPtr.Zero) return _hWndExcel;

                // NOTE: Don't use Process.GetCurrentProcess().MainWindowHandle; here,
                // it doesn't work when Excel is activated via COM, or when the add-in is installed.

                // NOTE: Careful not to call ExcelVersion here - might recurse causing StackOverflow
                // We try the Excel 2007+ case first.

                // Excel 2007+ - should get whole handle ... - we try this case first
                IntPtr hWnd = (IntPtr) (uint) (double) XlCall.Excel(XlCall.xlGetHwnd);
                // ... but this call is not reliable, and I'm not sure about 64-bit, 
                // so check if we actually got an Excel window, 
                // and even then only accept the result if there is more than the low word,
                // because many of the window functions seem to work even when passed only part of the window handle,
                // which could otherwise cause us to accept a partial handle.
                if (((uint)hWnd & 0xFFFF0000) != 0 && IsAnExcelWindow(hWnd))
                {
                    _hWndExcel = hWnd;
                    return _hWndExcel;
                }

                // Do a check based on the lo-Word - should work in all versions.
                ushort loWord = (ushort)hWnd;
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
            }, IntPtr.Zero);
            return hWnd;
        }

        static int _mainThreadId;
        internal static bool IsMainThread()
        {
            return Thread.CurrentThread.ManagedThreadId == _mainThreadId;
        }

        // Returns true if the cached _application reference is valid.
        // - someone might have called Marshal.ReleaseComObject, making this reference invalid.
        static bool IsApplicationOK()
        {
            if (_application == null) return false;
            try
            {
                _application.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, _application, null,  _enUsCulture);
                return true;
            }
            catch (Exception)
            {
                _application = null;
                return false;
            }
        }

        // CONSIDER: ThreadStatic not needed anymore - only cached and used on main thread anyway.
        // [ThreadStatic] 
        static object _application;
        static readonly CultureInfo _enUsCulture = new CultureInfo(1033);
		public static object Application
		{
			get
			{
                if (!IsMainThread())
                {
                    // Nothing cached - possibly being called on a different thread
                    // Just get from window and return
                    return GetApplication();
                }

                // Check whether we have a chached App and it is valid
                if (IsApplicationOK())
                {
                    return _application;
                }
                // There was a problem with the cached application.
                // Try to get one and remember  it.
                _application = GetApplication();
                return _application;
			}
            internal set
            {
                // Should only be set on the main thread.
                if (!IsMainThread()) throw new InvalidOperationException("Cached Application can only be set on the main thread.");
                _application = value;
            }
		}

        private static object GetApplication()
        {
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

        // CONSIDER: Might this be better?
        // return !XlCall.Excel(XlCall.xlfGetTool, 4, "Standard", 1);
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
			}, IntPtr.Zero);
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
                        string versionString = Regex.Match((string)versionObject, "^\\d+(\\.\\d+)?").Value;
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
                                                                             BindingFlags.GetProperty,
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
