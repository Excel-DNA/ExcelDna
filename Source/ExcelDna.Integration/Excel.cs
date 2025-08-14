//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelDna.Integration
{
    public static class ExcelDnaUtil
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

        private static bool checkForIllegalCrossThreadCalls;

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

            checkForIllegalCrossThreadCalls = Debugger.IsAttached;
        }

        internal static bool IsMainThread
        {
            get
            {
                return Thread.CurrentThread.ManagedThreadId == _mainManagedThreadId;
            }
        }

        public static int MainManagedThreadId
        {
            get { return _mainManagedThreadId; }
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
            EnumThreadWindows(_mainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
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
        static Microsoft.Office.Interop.Excel.Application _application;
        static readonly CultureInfo _enUsCulture = new CultureInfo(1033);
        public static object Application
        {
            get
            {
                bool isProtected;

                if (!IsMainThread)
                {
                    if (checkForIllegalCrossThreadCalls)
                    {
                        throw new InvalidOperationException(
                            "Cross-thread operation not valid: Application accessed from a thread other than the thread it was created on.");
                    }

                    // Nothing cached - possibly being called on a different thread
                    // Just get from window and return
                    return GetApplicationFromWindows(true, out isProtected);
                }

                // Check whether we have a cached App and it is valid
                if (IsApplicationOK())
                {
                    return _application;
                }

                // There was a problem with the cached application.
                _application = null;

                // Try to get one and remember  it.
                var application = GetApplication(true, out isProtected);
                if (!isProtected)
                {
                    // Only cache if it's not a protected Application object
                    _application = application;
                }
                return application;
            }
        }

        // This call might throw an access violation 
        // NOTE .NET5+: If this assembly is run under .NET 5+ we need to re-engineer this call to handle possible access violations outside the managed code,
        // or figure out the source and timing of safe vs dangerous calls.
        // (Also for CallPenHelper)
        [HandleProcessCorruptedStateExceptions]
        private static bool IsExcelApiAvailable()
        {
            try
            {
                object output;
                XlCall.XlReturn result = XlCall.TryExcel(XlCall.xlGetName, out output);
                if (result == XlCall.XlReturn.XlReturnFailed)
                {
                    // no plan for getting Application (we're probably on a different thread?)
                    // throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object.");
                    return false;
                }
            }
            catch (AccessViolationException)
            {
                // throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object. Excel may be shutting down", ave);
                return false;
            }
            return true;
        }

        private static Microsoft.Office.Interop.Excel.Application GetApplication(bool allowProtected, out bool isProtected)
        {
            // Don't cache the one we get from the Window, it keeps Excel alive! 
            // (?? Really ?? - Probably only when we're not on the main thread...)
            Microsoft.Office.Interop.Excel.Application application = GetApplicationFromWindows(allowProtected, out isProtected);
            if (application != null) return application;

            // DOCUMENT: Under some circumstances, the C API and Automation interfaces are not available.
            //  This happens when there is no Workbook open in Excel.
            // Now make workbook with VBA sheet, according to some Google post.

            // We try a (possible) test for whether we can call the C API.
            if (!IsExcelApiAvailable())
            {
                throw new InvalidOperationException("Excel API is unavailable - cannot retrieve Application object.");
            }

            return GetApplicationFromNewWorkbook(allowProtected, out isProtected);
        }

        private static Microsoft.Office.Interop.Excel.Application GetApplicationFromNewWorkbook(bool allowProtected, out bool isProtected)
        {
            // Create new workbook with the right stuff
            // Echo calls removed for Excel 2013 - this caused trouble in the Excel 2013 'browse' scenario.
            Microsoft.Office.Interop.Excel.Application application;
            bool isExcelPre15 = SafeIsExcelVersionPre15;
            if (isExcelPre15) XlCall.Excel(XlCall.xlcEcho, false);
            try
            {
                XlCall.Excel(XlCall.xlcNew, 5);
                XlCall.Excel(XlCall.xlcWorkbookInsert, 6);

                // Try again
                application = GetApplicationFromWindows(allowProtected, out isProtected);

                XlCall.Excel(XlCall.xlcFileClose, false);
            }
            catch
            {
                // Not expecting this ever - but be consistent about Try vs. exceptions
                application = null;
                isProtected = false;
            }
            finally
            {
                if (isExcelPre15) XlCall.Excel(XlCall.xlcEcho, true);
            }

            return application; // Might be null in a bad case, but we have no further ideas
        }

        // internal implementation that does not throw in the case where the C API is unavailable,
        // to improve QueueAsMacro reliability
        internal static object GetApplicationNotProtectedNoThrow()
        {
            if (!IsMainThread)
            {
                Debug.Fail("Must call GetApplicationNotProtectedNoThrow on the main thread");
                return null;
            }

            // Check for a good cached one
            if (IsApplicationOK())
            {
                return _application;
            }

            bool isProtected;
            var application = GetApplicationFromWindows(false, out isProtected);
            if (application != null && !isProtected)
            {
                // Only cache and use if not a protected application
                _application = application;
                return application;
            }

            // We try a (possible) test for whether we can call the C API.
            if (!IsExcelApiAvailable())
            {
                return null;
            }

            // We can call the C API - use it to make a new workbook and then get the Application through there
            application = GetApplicationFromNewWorkbook(false, out isProtected);
            if (application != null && isProtected)
            {
                // (We don't expect it to ever be protected in this case...)
                Debug.Fail("Unexpected protected Application from GetApplicationFromNewWorkbook");
                // Can't return this Application
                return null;
            }
            _application = application; // Still null due to unexpected failure, or else valid, not protected, and thus safe to cache
            return application;
        }

        static Microsoft.Office.Interop.Excel.Application GetApplicationFromWindows(bool allowProtected, out bool isProtected)
        {
            if (SafeIsExcelVersionPre15)
            {
                return GetApplicationFromWindow(WindowHandle, allowProtected, out isProtected);
            }

            return GetApplicationFromWindows15(allowProtected, out isProtected);
        }

        // Enumerate through all top-level windows of the main thread,
        // and for those of class XLMAIN, dig down by calling GetApplicationFromWindow.
        static Microsoft.Office.Interop.Excel.Application GetApplicationFromWindows15(bool allowProtected, out bool isProtected)
        {
            Microsoft.Office.Interop.Excel.Application application = null;
            StringBuilder buffer = new StringBuilder(256);
            bool localIsProtected = false;

            EnumThreadWindows(_mainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
            {
                // Check the window class
                if (IsAnExcelWindow(hWndEnum, buffer))
                {
                    application = GetApplicationFromWindow(hWndEnum, allowProtected, out localIsProtected);
                    if (application != null)
                    {
                        return false;	// Stop enumerating
                    }
                    return true;
                }
                return true;	// Continue enumerating
            }, IntPtr.Zero);
            isProtected = localIsProtected;
            return application; // May or may not be null
        }

        private static Microsoft.Office.Interop.Excel.Application GetApplicationFromWindow(IntPtr hWndMain, bool allowProtected, out bool isProtected)
        {
            // This is Andrew Whitechapel's plan for getting the Application object.
            // It does not work when there are no Workbooks open.
            Microsoft.Office.Interop.Excel.Application app = null;
            StringBuilder cname = new StringBuilder(256);
            bool localIsProtected = false;

            EnumChildWindows(hWndMain, delegate (IntPtr hWndEnum, IntPtr param)
            {
                // Check the window class
                GetClassNameW(hWndEnum, cname, cname.Capacity);
                if (cname.ToString() != "EXCEL7")
                    // Not a workbook window, continue enumerating
                    return true;

                IntPtr pUnk = IntPtr.Zero;
                int hr = AccessibleObjectFromWindow(hWndEnum, OBJID_NATIVEOM, IID_IDispatchBytes, ref pUnk);
                if (hr != 0)
                {
                    // Window does not implement the IID, continue enumerating
                    return true;
                }

                // Marshal to .NET, then call .Application
                object obj = Marshal.GetObjectForIUnknown(pUnk);
                Marshal.Release(pUnk);

                try
                {
                    if (ComInterop.DispatchHelper.HasProperty(obj, "Application"))
                    {
                        app = (Microsoft.Office.Interop.Excel.Application)obj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, obj, null, _enUsCulture);
                    }
                    else
                    {
                        // In some cases - always when Excel only a workbook open in Protected Mode when this code runs - 
                        // we get a ProtectedViewWindow, which has no Application property, but then we can do .Workbook.Application
                        if (allowProtected)
                        {
                            try
                            {
                                object workbook = obj.GetType().InvokeMember("Workbook", BindingFlags.GetProperty, null, obj, null, _enUsCulture);
                                app = (Microsoft.Office.Interop.Excel.Application)workbook.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, workbook, null, _enUsCulture);

                                // WARNING: The Application object returning from here can be problematic:
                                //          * It is a "sandbox" view of the Application that cannot Run macros or change workbooks
                                //          * It will die after the protected view closes
                                //          We return it, but should not cache it
                                localIsProtected = true;
                            }
                            catch
                            {
                                // Otherwise we fail - this way the higher-level call will open up a regular workbook and try again
                                Debug.Fail("Unexpected exception when getting Application");
                            }
                        }
                        else
                        {
                            // We don't want to continue enumeration in this case - another window might give us an Application, but it will be a protected one anyway
                            localIsProtected = true;
                        }
                    }
                }
                catch
                {
                    // Otherwise we fail - this way the higher-level call will open up a regular workbook and try again
                    Debug.Fail("Unexpected exception when getting Application");
                }
                finally
                {
                    Marshal.ReleaseComObject(obj);
                }

                // Continue enumeration? Only if the app is not yet found and protected flag not set.
                return (app == null) && !localIsProtected;
            }, IntPtr.Zero);
            isProtected = localIsProtected;
            return app;
        }

        // Returns true if the cached _application reference is valid.
        // - someone might have called Marshal.ReleaseComObject, making this reference invalid.
        static bool IsApplicationOK()
        {
            if (_application == null) return false;
            try
            {
                // TODO: Can we turn this into a delegate somehow - for performance and to avoid the exception.
                //       I'm not sure how to pull out the Property if it's name might be based on the Culture.
                //       One way is to get the dispid, and call through IDispatch.Invoke, but that's a lot of work...
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
            EnumThreadWindows(_mainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
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

        struct FuncWizChild
        {
            public int ScrollBar;
            public int EDTBX;
        };

        static bool IsFunctionWizardWindow(IntPtr hWnd, StringBuilder buffer)
        {
            buffer.Length = 0;
            // Check the window class
            if (GetClassNameW(hWnd, buffer, buffer.Capacity) == 0)
                return false;
            if (!buffer.ToString().StartsWith("bosa_sdm_XL"))
                return false;

            FuncWizChild child = new FuncWizChild { ScrollBar = 0, EDTBX = 0 };
            EnumChildWindows(hWnd, delegate (IntPtr hWndEnum, IntPtr param)
            {
                buffer.Length = 0;
                if (GetClassNameW(hWndEnum, buffer, buffer.Capacity) == 0)
                    return false;

                string title = buffer.ToString();
                if (title.Equals("EDTBX"))
                    child.EDTBX++;
                else if (title.Equals("ScrollBar"))
                    child.ScrollBar++;
                else
                    return false;

                return true;
            }, IntPtr.Zero);

            if (child.ScrollBar == 1 && child.EDTBX == 5)
                return true;

            return false;
        }
        #endregion

        #region IsInFormualEditMode
        public static bool IsInFormulaEditMode()
        {
            if (!IsMainThread)
                throw new InvalidOperationException("IsInFormulaEditMode can only be called from the main thread.");

            return RunMacroSynchronization.IsInFormulaEditMode();
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
            EnumThreadWindows(_mainNativeThreadId, delegate (IntPtr hWndEnum, IntPtr param)
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

        public static FileInfo XllPathInfo
        {
            get
            {
                return DnaLibrary.XllPathInfo;
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
        static readonly ExcelLimits _xlLimits = new ExcelLimits
        {
            MaxRows = 1048576,
            MaxColumns = 16384,
            MaxArguments = 256,
            MaxStringLength = 32767
        };
        public static ExcelLimits ExcelLimits => _xlLimits;
        #endregion

        #region SupportsDynamicArrays
        static bool? _supportsDynamicArrays;
        public static bool SupportsDynamicArrays
        {
            get
            {
                if (!_supportsDynamicArrays.HasValue)
                {
                    object result;
                    var returnValue = XlCall.TryExcel(614, out result, new object[] { 1 }, new object[] { true }); // 614 means FILTER
                    // Now examine returnValue, which should be of type XlReturn � it will presumably be XlReturn.XlReturnSuccess for Dynamic Array Excel, otherwise XlReturn.XlReturnFailed or similar for non-DA Excel.
                    _supportsDynamicArrays = (returnValue == XlCall.XlReturn.XlReturnSuccess);
                }
                return _supportsDynamicArrays.Value;
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
