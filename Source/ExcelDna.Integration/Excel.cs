//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace ExcelDna.Integration
{
    public class ExcelDnaUtil
    {
        private delegate bool EnumWindowsCallback(IntPtr hwnd, /*ref*/ IntPtr param);
        private delegate bool EnumThreadWindowsCallback(IntPtr hwnd, /*ref*/ IntPtr param);

        [DllImport("user32.dll")]
        private static extern int EnumWindows(EnumWindowsCallback callback, /*ref*/ IntPtr param);
        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadWindowsCallback callback, /*ref*/ IntPtr param);
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
        // Use the overload that's convenient when we don't need the ProcessId - pass IntPtr.Zero for the second parameter
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, /*out uint */ IntPtr refProcessId);
        [DllImport("Kernel32")]
        private static extern uint GetCurrentThreadId();

        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly byte[] IID_IDispatchBytes = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();


        // Some static state, set when ExcelDna.Integration is initialized.
        static uint _mainNativeThreadId;
        static int _mainManagedThreadId;
        static IntPtr _mainWindowHandle;

        internal static void Initialize()
        {
            // NOTE: Sequence here is important - Getting the Window Handle sometimes uses the _mainNativeThreadId
            _mainManagedThreadId = Thread.CurrentThread.ManagedThreadId;
            _mainNativeThreadId = GetCurrentThreadId();

            // Under Excel 2013 we have a problem, the window is not stable.
            // We get the current main window here, and deal with changes later.
            _mainWindowHandle = GetWindowHandleApi();
            // Extra attempt via window enumeration if the Api approach failed.
            if (_mainWindowHandle == IntPtr.Zero)
                _mainWindowHandle = GetWindowHandleThread();
        }

        internal static bool IsMainThread
        {
            get
            {
                return Thread.CurrentThread.ManagedThreadId == _mainManagedThreadId;
            }
        }

        #region Get Window Handle
        // NOTE: Careful not to call ExcelVersion here - might recurse causing StackOverflow

        // NOTE: Previously we cached the value as an optimization (because the IsAnExcelWindow check below may be slow)
        //       But under Excel 2013 the window handle is not stable - it's merely the most recently active of the SDI main windows.
        //       Now that we (a) cache the Application object on the main thread, 
        //       and (b) can run everything on the main thread with ExcelAsyncUtil, this is not really a problem.

        // NOTE: Don't use Process.GetCurrentProcess().MainWindowHandle; here,
        // it doesn't work when Excel is activated via COM, or when the add-in is installed.

        public static IntPtr WindowHandle
        {
            get
            {
                if (SafeIsExcelVersionPre15) return _mainWindowHandle;

                // Under Excel 2013, the window handles change according to the active workbook.
                return GetWindowHandle15();
            }
        }

        static IntPtr GetWindowHandle15()
        {
            IntPtr hWnd = IntPtr.Zero;
            if (IsMainThread)
            {
                // We're on the main thread, and might already have an application object.
                if (_application != null)
                {
                    // This is the typical case.
                    // Get it from there - probably safest
                    // TODO: Can we turn this into a delegate somehow...(for performance)?
                    hWnd = (IntPtr)(int)_application.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, _application, null, _enUsCulture);
                    if (IsWindowOfThisExcel(hWnd)) return hWnd;
                }

                // We're on the main thread, try the C API directly
                hWnd = GetWindowHandleApi();
                if (hWnd != IntPtr.Zero && IsWindowOfThisExcel(hWnd)) return hWnd;
            }

            StringBuilder buffer = new StringBuilder(256);
            // If the main window handle stored in Initialization is still valid, use that.
            if (IsAnExcelWindow(_mainWindowHandle, buffer))
                return _mainWindowHandle;

            hWnd = GetWindowHandleThread();

            if (hWnd != IntPtr.Zero)
            {
                // Change the MainWindowHandle - the previous one is no longer valid
                _mainWindowHandle = hWnd;
                return hWnd;
            }

            // Give up
            throw new InvalidOperationException("Window handle cannot be retrieved at this time.");
        }

        // Get the window handle using the C API - should work on all Excel versions,
        // but not in funny circumstances, like cluster or regsvr32...
        static IntPtr GetWindowHandleApi()
        {
            // Should not throw exception
            // (we expect the C API call to fail when running on Cluster or during automation)
            try
            {
                double apiHwnd = (double)XlCall.Excel(XlCall.xlGetHwnd);

                // I have no idea how Excel converts the window handle into a double.
                // Under 32-bit I've had reported values > Int32.MaxValue and negative values...

                IntPtr hWnd;
                if (apiHwnd < 0)
                    hWnd = (IntPtr)(int)apiHwnd;
                else
                    hWnd = (IntPtr)(uint)apiHwnd;

                if (IsWindowOfThisExcel(hWnd)) return hWnd;

                if (hWnd != IntPtr.Zero)
                {
                    // We might have a partial handle... try to find a window that matches.
                    // Do a check based on the lo-Word - should work in all versions.
                    ushort loWord = (ushort)hWnd;
                    hWnd = FindAnExcelWindowWithLoWord(loWord);
                    if (IsWindowOfThisExcel(hWnd)) return hWnd;
                }
            }
            catch (Exception e)
            {
                Debug.Write("GetWindowHandleApi error - " + e);
                // Ignore errors
            }

            // This is pretty bad - caller needs to try another approach.
            return IntPtr.Zero;
        }

        // Tries to get the window handle by enumerating thread windows of the main thread, 
        // and accepting any XLMAIN window.
        // Returns Zero if that fails.
        static IntPtr GetWindowHandleThread()
        {
            IntPtr hWnd = IntPtr.Zero;

            StringBuilder buffer = new StringBuilder(255);
            EnumThreadWindows(_mainNativeThreadId, delegate(IntPtr hWndEnum, IntPtr param)
            {
                if (IsAnExcelWindow(hWndEnum, buffer))
                {
                    hWnd = hWndEnum;
                    return false;	// Stop enumerating
                }
                return true;	// Continue enumerating
            }, IntPtr.Zero);

            // hWnd might still be Zero...?
            return hWnd;
        }

        // Checks whether a handle is a valid window handle on our main thread.
        static bool IsWindowOfThisExcel(IntPtr hWnd)
        {
            uint threadId = GetWindowThreadProcessId(hWnd, IntPtr.Zero);
            // returns 0 if the window handle is not valid.
            return threadId == _mainNativeThreadId;
        }
        #endregion

        #region Get Application COM Object
        // CONSIDER: ThreadStatic not needed anymore - only cached and used on main thread anyway.
        // [ThreadStatic] 
        static object _application;
        static readonly CultureInfo _enUsCulture = new CultureInfo(1033);
        public static object Application
        {
            get
            {
                if (!IsMainThread)
                {
                    // Nothing cached - possibly being called on a different thread
                    // Just get from window and return
                    return GetApplicationFromWindows();
                }

                // Check whether we have a cached App and it is valid
                if (IsApplicationOK())
                {
                    return _application;
                }
                // There was a problem with the cached application.
                // Try to get one and remember  it.
                _application = GetApplication();
                return _application;
            }
        }

        // This call might throw an access violation 
        // .NET40: If this assembly is compiled for .NET 4, add this attribute to get the expected behaviour.
        // (Also for CallPenHelper)
        // [HandleProcessCorruptedStateExceptions]
        private static void CheckExcelApiAvailable()
        {
            try
            {
                object output;
                XlCall.XlReturn result = XlCall.TryExcel(XlCall.xlGetName, out output);
                if (result == XlCall.XlReturn.XlReturnFailed)
                {
                    // no plan for getting Application (we're probably on a different thread?)
                    throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object.");
                }
            }
            catch (AccessViolationException ave)
            {
                throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object. Excel may be shutting down", ave);
            }
        }

        private static object GetApplication()
        {
            // Don't cache the one we get from the Window, it keeps Excel alive! 
            // (?? Really ?? - Probably only when we're not on the main thread...)
            object application = GetApplicationFromWindows();
            if (application != null) return application;
            
            // DOCUMENT: Under some circumstances, the C API and Automation interfaces are not available.
            //  This happens when there is no Workbook open in Excel.
            // Now make workbook with VBA sheet, according to some Google post.

            // We try a (possible) test for whether we can call the C API.
            CheckExcelApiAvailable();

            // Create new workbook with the right stuff
            // Echo calls removed for Excel 2013 - this caused trouble in the Excel 2013 'browse' scenario.
            bool isExcelPre15 = SafeIsExcelVersionPre15;
            if (isExcelPre15) XlCall.Excel(XlCall.xlcEcho, false);

            XlCall.Excel(XlCall.xlcNew, 5);
            XlCall.Excel(XlCall.xlcWorkbookInsert, 6);

            // Try again
            application = GetApplicationFromWindows();

            // Clean up
            XlCall.Excel(XlCall.xlcFileClose, false);
            if (isExcelPre15) XlCall.Excel(XlCall.xlcEcho, true);

            if (application != null) return application;
            
            // This is really bad - throwing an exception ...
            throw new InvalidOperationException("Excel Application object could not be retrieved.");
        }

        static object GetApplicationFromWindows()
        {
            if (SafeIsExcelVersionPre15)
            {
                return GetApplicationFromWindow(WindowHandle);
            }

            return GetApplicationFromWindows15();
        }

        // Enumerate through all top-level windows of the main thread,
        // and for those of class XLMAIN, dig down by calling GetApplicationFromWindow.
        static object GetApplicationFromWindows15()
        {
            object application = null;
            StringBuilder buffer = new StringBuilder(256);
            EnumThreadWindows(_mainNativeThreadId, delegate(IntPtr hWndEnum, IntPtr param)
            {
                // Check the window class
                if (IsAnExcelWindow(hWndEnum, buffer))
                {
                    application = GetApplicationFromWindow(hWndEnum);
                    if (application != null)
                    {
                        return false;	// Stop enumerating
                    }
                    return true;
                }
                return true;	// Continue enumerating
            }, IntPtr.Zero);
            return application; // May or may not be null
        }

        private static object GetApplicationFromWindow(IntPtr hWndMain)
        {
            // This is Andrew Whitechapel's plan for getting the Application object.
            // It does not work when there are no Workbooks open.
            IntPtr hWndChild = IntPtr.Zero;
            StringBuilder cname = new StringBuilder(256);
            EnumChildWindows(hWndMain, delegate(IntPtr hWndEnum, IntPtr param)
            {
                // Check the window class
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
                    // Marshal to .NET, then call .Application
                    object obj = Marshal.GetObjectForIUnknown(pUnk);
                    Marshal.Release(pUnk);

                    object app = obj.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, obj, null, new CultureInfo(1033));
                    Marshal.ReleaseComObject(obj);

                    //   object ver = app.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, app, null);
                    return app;
                }
            }
            return null;
        }

        // Returns true if the cached _application reference is valid.
        // - someone might have called Marshal.ReleaseComObject, making this reference invalid.
        static bool IsApplicationOK()
        {
            if (_application == null) return false;
            try
            {
                // TODO: Can we turn this into a delegate somehow - for performance.
                _application.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, _application, null, _enUsCulture);
                return true;
            }
            catch (Exception)
            {
                _application = null;
                return false;
            }
        }
        #endregion

        #region IsInFunctionWizard

        // CONSIDER: Might this be better?
        // return !XlCall.Excel(XlCall.xlfGetTool, 4, "Standard", 1);
        // (I think not - apparently it doesn't work right under Excel 2013, 
        //  but it can easily be called by the user in their helper)
        // We enumerate these by getting the (native) thread id, from any WindowHandle,
        // then enumerating non-child windows for this thread.
        public static bool IsInFunctionWizard()
        {
            // TODO: Handle the Find and Replace dialog
            //       for international versions.
            StringBuilder buffer = new StringBuilder(256);
            bool inFunctionWizard = false;
            EnumThreadWindows(_mainNativeThreadId, delegate(IntPtr hWndEnum, IntPtr param)
            {
                if (IsFunctionWizardWindow(hWndEnum, buffer))
                {
                    inFunctionWizard = true;
                    return false; // Stop enumerating
                }
                return true;	// Continue enumerating
            }, IntPtr.Zero);
            return inFunctionWizard;
        }

        static bool IsFunctionWizardWindow(IntPtr hWnd, StringBuilder buffer)
        {
            buffer.Length = 0;
            // Check the window class
            GetClassNameW(hWnd, buffer, buffer.Capacity);
            if (!buffer.ToString().StartsWith("bosa_sdm_XL")) 
                return false;

            buffer.Length = 0;
            GetWindowTextW(hWnd, buffer, buffer.Capacity);
            string title = buffer.ToString();
            // Another window that has been reported as causing issue has title "Collect and Paste 2.0"
            if (title.Contains("Replace") || title.Contains("Paste") || title.Contains("Recovery"))
                return false;

            return true;
        }
        #endregion

        #region Version Helpers
        // This version is used internally - it seems a bit safer than the API calls.
        private static FileVersionInfo _excelExecutableInfo = null;
        internal static FileVersionInfo ExcelExecutableInfo
        {
            get
            {
                if (_excelExecutableInfo == null)
                {
                    ProcessModule excel = Process.GetCurrentProcess().MainModule;
                    _excelExecutableInfo = FileVersionInfo.GetVersionInfo(excel.FileName);
                }
                return _excelExecutableInfo;
            }
        }

        // One of our many, many version helpers...
        // This is safer than ExcelVersion since we can call it from initialization 
        // (it does not use COM Application object) and we can call it from an IsMacroType=false function.
        internal static bool SafeIsExcelVersionPre15
        {
            get
            {
                return ExcelExecutableInfo.FileMajorPart < 15;
            }
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
        #endregion

        #region Window Helpers

        // Check if hWnd refers to a Window of class "XLMAIN" indicating an Excel top-level window.
        static bool IsAnExcelWindow(IntPtr hWnd, StringBuilder buffer)
        {
            buffer.Length = 0;
            GetClassNameW(hWnd, buffer, buffer.Capacity);
            return buffer.ToString() == "XLMAIN";
        }

        // Try to find an Excel window with window handle that matches the passed lo word.
        static IntPtr FindAnExcelWindowWithLoWord(ushort hWndLoWord)
        {
            IntPtr hWnd = IntPtr.Zero;
            StringBuilder buffer = new StringBuilder(256);
            EnumThreadWindows(_mainNativeThreadId, delegate(IntPtr hWndEnum, IntPtr param)
            {
                // Check the loWord
                if (((uint)hWndEnum & 0x0000FFFF) == (uint)hWndLoWord &&
                    IsAnExcelWindow(hWndEnum, buffer))
                {
                    hWnd = hWndEnum;
                    return false;  // Stop enumerating
                }
                return true;  // Continue enumerating
            }, IntPtr.Zero);
            return hWnd;
        }
        #endregion

        #region Path Helpers
        // Public access to the XllPath, safe in any context
        public static string XllPath
        {
            get
            {
                return DnaLibrary.XllPath;
            }
        }

        static readonly Guid _excelDnaNamespaceGuid = new Guid("{306D016E-CCE8-4861-9DA1-51A27CBE341A}");
        internal static Guid XllGuid { get { return GuidFromXllPath(XllPath); } }

        // Return a stable Guid from the xll path - used for COM registration and helper functions
        // Uses the .ToUpperInvariant() of the path name.
        // CONSIDER: Should we use only the file name, not the full path (the path used might not be canonical)
        //           I'm still a bit unsure about having different add-ins with the same file name.
        internal static Guid GuidFromXllPath(string path)
        {
            return GuidUtility.Create(_excelDnaNamespaceGuid, path.ToUpperInvariant());
        }
        #endregion

        #region ExcelLimits
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
                        _xlLimits.MaxArguments = 256;
                        _xlLimits.MaxStringLength = 32767;
                    }
                }
                return _xlLimits;
            }
        }
        #endregion
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
