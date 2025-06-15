﻿//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Loader.Logging;

namespace ExcelDna.Loader
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;
    internal delegate void fn_void_double(double dValue);
    internal delegate short fn_short_void();
    internal delegate void fn_void_intptr(IntPtr intPtr);
    internal delegate void fn_void_void();
    internal delegate IntPtr fn_intptr_intptr(IntPtr intPtr);
    internal delegate HRESULT fn_hresult_void();
    internal delegate HRESULT fn_get_class_object(CLSID rclsid, IID riid, out IntPtr ppunk);

    // CAUTION: This struct is also defined in the unmanaged loader.
    internal struct XlAddInExportInfo
    {
#pragma warning disable 0649 // Field 'field' is never assigned to, and will always have its default value 'value'
        internal Int32 ExportInfoVersion; // Must be 10 for this version
        internal Int32 AppDomainId; // Id of the Sandbox AppDomain where the add-in runs.
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoOpen;
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoClose;
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoRemove;
        internal IntPtr /* PFN_VOID_LPXLOPER12 */       pXlAutoFree12;
        internal IntPtr /* PFN_PFNEXCEL12 */            pSetExcel12EntryPt;
        internal IntPtr /* PFN_HRESULT_VOID */          pDllRegisterServer;
        internal IntPtr /* PFN_HRESULT_VOID */          pDllUnregisterServer;
        internal IntPtr /* PFN_GET_CLASS_OBJECT */      pDllGetClassObject;
        internal IntPtr /* PFN_HRESULT_VOID */          pDllCanUnloadNow;
        internal IntPtr /* PFN_VOID_DOUBLE */           pSyncMacro;
        internal IntPtr /* PFN_LPXLOPER12_LPXLOPER12 */ pRegistrationInfo;
        internal IntPtr /* PFN_VOID_VOID */             pCalculationCanceled;
        internal IntPtr /* PFN_VOID_VOID */             pCalculationEnded;
        internal IntPtr /* PFN_LPENHELPER */            pLPenHelper;
        internal Int32 ThunkTableLength;  // Must be EXPORT_COUNT
        internal IntPtr /*PFN*/ ThunkTable; // Actually (PFN ThunkTable[EXPORT_COUNT])
#pragma warning restore 0649
    };

    // CAUTION: This type is loaded by reflection from the unmanaged loader.
    public unsafe static class XlAddIn
    {
        static int thunkTableLength;
        static IntPtr thunkTable;

        // Passed in from unmanaged code during initialization 
        internal static IntPtr hModuleXll;

        // Consider: When does this become an interface. etc
        internal static string PathXll { get; private set; }
        internal static string TempDirPath { get; private set; }
        internal static Func<string, int, byte[]> GetResourceBytes;  // Passed in from Loader
        internal static Func<string, Assembly> LoadAssemblyFromPath;  // Passed in from Loader
        internal static Func<byte[], byte[], Assembly> LoadAssemblyFromBytes;  // Passed in from Loader
        internal static Action<TraceSource> SetIntegrationTraceSource;  // Passed in from Loader

        internal delegate int LPenHelperDelegate(int wCode, ref XlCall.FmlaInfo fmlaInfo);
        internal static LPenHelperDelegate LPenHelper;

        static int xlCallVersion;
        internal static int XlCallVersion { get { return xlCallVersion; } }

        static bool _initialized = false;
        static bool _opened = false;

        static List<GCHandle> fnHandles = new List<GCHandle>();

        #region Initialization

        public static bool Initialize(IntPtr xlAddInExportInfoAddress, IntPtr hModuleXll, string pathXll, string tempDirPath,
                                             Func<string, int, byte[]> getResourceBytes,
                                             Func<string, Assembly> loadAssemblyFromPath,
                                             Func<byte[], byte[], Assembly> loadAssemblyFromBytes,
                                             Action<TraceSource> setIntegrationTraceSource)
        {
            XlAddIn.hModuleXll = hModuleXll;
            XlAddIn.PathXll = pathXll;
            XlAddIn.TempDirPath = Path.Combine(tempDirPath ?? Path.GetTempPath(), "ExcelDna.Loader");
            XlAddIn.GetResourceBytes = getResourceBytes;
            XlAddIn.LoadAssemblyFromPath = loadAssemblyFromPath;
            XlAddIn.LoadAssemblyFromBytes = loadAssemblyFromBytes;
            XlAddIn.SetIntegrationTraceSource = setIntegrationTraceSource;

            // NOTE: Too early for logging - the TraceSource in ExcelDna.Integration has not been initialized yet.
            Debug.Print("In sandbox AppDomain with Id: {0}, running on thread: {1}", AppDomain.CurrentDomain.Id, Thread.CurrentThread.ManagedThreadId);
            Debug.Assert(xlAddInExportInfoAddress != IntPtr.Zero, "InitializationInfo address is null");
            Debug.Print("InitializationInfo address: 0x{0:x8}", xlAddInExportInfoAddress);

            XlAddInExportInfo* pXlAddInExportInfo = (XlAddInExportInfo*)xlAddInExportInfoAddress;
            int exportInfoVersion = pXlAddInExportInfo->ExportInfoVersion;
            if (exportInfoVersion != 10)
            {
                Debug.Print("ExportInfoVersion '{0}' not supported", exportInfoVersion);
                return false;
            }

            fn_short_void fnXlAutoOpen = (fn_short_void)XlAutoOpen;
            fnHandles.Add(GCHandle.Alloc(fnXlAutoOpen));
            pXlAddInExportInfo->pXlAutoOpen = Marshal.GetFunctionPointerForDelegate(fnXlAutoOpen);

            fn_short_void fnXlAutoClose = (fn_short_void)XlAutoClose;
            fnHandles.Add(GCHandle.Alloc(fnXlAutoClose));
            pXlAddInExportInfo->pXlAutoClose = Marshal.GetFunctionPointerForDelegate(fnXlAutoClose);

            fn_short_void fnXlAutoRemove = (fn_short_void)XlAutoRemove;
            fnHandles.Add(GCHandle.Alloc(fnXlAutoRemove));
            pXlAddInExportInfo->pXlAutoRemove = Marshal.GetFunctionPointerForDelegate(fnXlAutoRemove);

            fn_void_intptr fnXlAutoFree12 = (fn_void_intptr)XlAutoFree12;
            fnHandles.Add(GCHandle.Alloc(fnXlAutoFree12));
            pXlAddInExportInfo->pXlAutoFree12 = Marshal.GetFunctionPointerForDelegate(fnXlAutoFree12);

            fn_void_intptr fnSetExcel12EntryPt = (fn_void_intptr)XlCallImpl.SetExcel12EntryPt;
            fnHandles.Add(GCHandle.Alloc(fnSetExcel12EntryPt));
            pXlAddInExportInfo->pSetExcel12EntryPt = Marshal.GetFunctionPointerForDelegate(fnSetExcel12EntryPt);

            fn_hresult_void fnDllRegisterServer = (fn_hresult_void)DllRegisterServer;
            fnHandles.Add(GCHandle.Alloc(fnDllRegisterServer));
            pXlAddInExportInfo->pDllRegisterServer = Marshal.GetFunctionPointerForDelegate(fnDllRegisterServer);

            fn_hresult_void fnDllUnregisterServer = (fn_hresult_void)DllUnregisterServer;
            fnHandles.Add(GCHandle.Alloc(fnDllUnregisterServer));
            pXlAddInExportInfo->pDllUnregisterServer = Marshal.GetFunctionPointerForDelegate(fnDllUnregisterServer);

            fn_get_class_object fnDllGetClassObject = (fn_get_class_object)DllGetClassObject;
            fnHandles.Add(GCHandle.Alloc(fnDllGetClassObject));
            pXlAddInExportInfo->pDllGetClassObject = Marshal.GetFunctionPointerForDelegate(fnDllGetClassObject);

            fn_hresult_void fnDllCanUnloadNow = (fn_hresult_void)DllCanUnloadNow;
            fnHandles.Add(GCHandle.Alloc(fnDllCanUnloadNow));
            pXlAddInExportInfo->pDllCanUnloadNow = Marshal.GetFunctionPointerForDelegate(fnDllCanUnloadNow);

            fn_void_double fnSyncMacro = (fn_void_double)SyncMacro;
            fnHandles.Add(GCHandle.Alloc(fnSyncMacro));
            pXlAddInExportInfo->pSyncMacro = Marshal.GetFunctionPointerForDelegate(fnSyncMacro);

            fn_intptr_intptr fnRegistrationInfo = (fn_intptr_intptr)RegistrationInfo;
            fnHandles.Add(GCHandle.Alloc(fnRegistrationInfo));
            pXlAddInExportInfo->pRegistrationInfo = Marshal.GetFunctionPointerForDelegate(fnRegistrationInfo);

            fn_void_void fnCalculationCanceled = (fn_void_void)CalculationCanceled;
            fnHandles.Add(GCHandle.Alloc(fnCalculationCanceled));
            pXlAddInExportInfo->pCalculationCanceled = Marshal.GetFunctionPointerForDelegate(fnCalculationCanceled);

            fn_void_void fnCalculationEnded = (fn_void_void)CalculationEnded;
            fnHandles.Add(GCHandle.Alloc(fnCalculationEnded));
            pXlAddInExportInfo->pCalculationEnded = Marshal.GetFunctionPointerForDelegate(fnCalculationEnded);

            LPenHelper = (LPenHelperDelegate)Marshal.GetDelegateForFunctionPointer(pXlAddInExportInfo->pLPenHelper, typeof(LPenHelperDelegate));

            // Thunk table for registered functions
            thunkTableLength = pXlAddInExportInfo->ThunkTableLength;
            thunkTable = pXlAddInExportInfo->ThunkTable;

            // This is the place where we call into Excel - this causes SecurityPermission exception
            // when run from VSTO. I don't know how to deal with this problem yet.
            // TODO: Learn more about the special security stuff in VSTO.
            //       See ExecutionContext.SuppressFlow links below.
            try
            {
                XlAddIn.xlCallVersion = XlCallImpl.XLCallVer() / 256;
            }
            catch (DllNotFoundException)
            {
                // This is expected if we are running under HPC or Regsvr32.
                Debug.Print("XlCall library not found - probably running in HPC host or Regsvr32.exe");

                // For the HPC support, I ignore error here and just assume we are under new Excel.
                // This will cause the common error here to get pushed to later ...
                XlAddIn.xlCallVersion = 12;
                // return false;
            }
            catch (Exception e)
            {
                Debug.Print("XlAddIn: XLCallVer Error: {0}", e);

                // CONSIDER: Is this right / needed - I'm not actually sure what happens under HPC host, 
                // so I'll leave this case in here too.?
                XlAddIn.xlCallVersion = 12;
                // return false;
            }

            try
            {
                IntegrationLoader.LoadIntegration();
            }
            catch (InvalidOperationException ioe)
            {
                Debug.Print("XlAddIn: Initialize Error - Invalid ExcelIntegration version detected: {0}", ioe);
                return false;
            }
            catch (Exception e)
            {
                Debug.Print("XlAddIn: Initialize Error:", e);
                return false;
            }

            // File.AppendAllText(Path.ChangeExtension(pathXll, ".log"), string.Format("{0:u} XlAddIn.Initialize OK\r\n", DateTime.Now));

            return true;
        }

        internal static unsafe void SetJump(int fi, IntPtr pfn)
        {
            if (fi >= 0 && fi < thunkTableLength)
            {
                void** pThunkTable = (void**)(thunkTable);
                pThunkTable[fi] = (void*)pfn;
            }
        }

        private static void InitializeIntegration()
        {
            if (!_initialized)
            {
                ExcelIntegration.Initialize(PathXll);
                TraceLogger.IntegrationTraceSource = ExcelIntegration.GetIntegrationTraceSource();
                SetIntegrationTraceSource(TraceLogger.IntegrationTraceSource);
                _initialized = true;
            }
        }

        private static void DeInitializeIntegration()
        {
            if (_initialized)
            {
                if (_opened)
                {
                    ExcelIntegration.DnaLibraryAutoClose();
                    XlRegistration.UnregisterMethods();
                }
                TraceLogger.IntegrationTraceSource = null;
                SetIntegrationTraceSource(null);
                ExcelIntegration.DeInitialize();
                _initialized = false;
                _opened = false;
            }

            fnHandles.ForEach(i => i.Free());
            fnHandles.Clear();
        }
        #endregion

        #region Managed Xlxxxx functions
        internal static short XlAutoOpen()
        {
            Debug.Print("XlAddIn.XlAutoOpen - AppDomain Id: " + AppDomain.CurrentDomain.Id + " (Default: " + AppDomain.CurrentDomain.IsDefaultAppDomain() + ")");
            short result = 0;
            try
            {
                if (_opened)
                {
                    DeInitializeIntegration();
                }
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignore*/ , true, "Registering library " + PathXll);
                InitializeIntegration();
                Logger.Initialization.Verbose("In XlAddIn.XlAutoOpen");

                // v. 30 - moved the setting of _opened before calling AutoOpen, 
                // so that checking in DeInitializeIntegration does not prevent AutoOpen - unloading via xlAutoRemove from working.
                _opened = true;

                // InitializeIntegration has loaded the DnaLibrary
                ExcelIntegration.DnaLibraryAutoOpen();

                result = 1; // All is OK

                // Clear the status bar message
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignored*/ , false);
                // Debug.Print("Clear status bar message result: " + xlCallResult);
            }
            catch (Exception e)
            {
                // Can't use logging, xlcAlert and xlcMessage with length >220 here
                string message = string.Format("ExcelDna add-in InitializeIntegration failed - {1} - {0}", PathXll, e.Message);
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignore*/ , true, message.Substring(0, Math.Min(message.Length, 220)));
                result = 0;
            }
            return result;
        }

        internal static short XlAutoClose()
        {
            short result = 0;
            try
            {
                Logger.Initialization.Verbose("In XlAddIn.XlAutoClose");
                // This also gets called when workbook starts closing - and can still be cancelled
                result = 1; // 0 if problems ?
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "XlAddIn.XlAutoClose error");
            }

            return result;
        }

        // No longer exported (or called) from native loader.
        internal static short XlAutoAdd()
        {
            // AutoOpen will get called too, where we set everything up ...
            short result = 0;
            try
            {
                Logger.Initialization.Verbose("In XlAddIn.XlAutoAdd");
                result = 1;
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "XlAddIn.XlAutoAdd error");
            }

            return result;
        }

        internal static short XlAutoRemove()
        {
            short result = 0;
            try
            {
                Logger.Initialization.Verbose("In XlAddIn.XlAutoRemove");
                // Apparently better if called here, 
                // so I try to, but make it safe to call again.
                DeInitializeIntegration();
                return 1; // 0 if problems ?
            }
            catch (Exception e)
            {
                Logger.Initialization.Error(e, "XlAddIn.XlAutoRemove error");
            }

            return result;
        }

        internal static void XlAutoFree12(IntPtr pXloper12)
        {
            // CONSIDER: This might be improved....
            // Another option would be to have the Com memory allocator run in unmanaged code.
            // Right now I think this is OK, and easiest from where I'm coming.
            // This function can only be called after a return from a user function.
            // I just free all the possibly big memory allocations.

            XlDirectMarshal.FreeMemory();
        }
        #endregion

        #region Com Server exports
        internal static HRESULT DllRegisterServer()
        {
            InitializeIntegration();
            return ExcelIntegration.DllRegisterServer();
        }

        internal static HRESULT DllUnregisterServer()
        {
            InitializeIntegration();
            return ExcelIntegration.DllUnregisterServer();
        }

        internal static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        {
            Debug.Print("DllGetClassObject entered - calling InitializeIntegration.");
            HRESULT result;
            InitializeIntegration();
            Logger.Initialization.Verbose("In DllGetClassObject");
            result = ExcelIntegration.DllGetClassObject(clsid, iid, out ppunk);
            return result;
        }

        internal static HRESULT DllCanUnloadNow()
        {
            InitializeIntegration();
            return ExcelIntegration.DllCanUnloadNow();
        }
        #endregion

        #region Extensions support
        internal static void SyncMacro(double dValue)
        {
            if (_initialized)
                ExcelIntegration.SyncMacro(dValue);
        }

        internal static IntPtr RegistrationInfo(IntPtr pParam)
        {
            if (!_initialized)
            {
                return IntPtr.Zero;
            }

            object param = XlMarshalContext.ObjectParam(pParam);
            object regInfo = XlRegistration.GetRegistrationInfo(param);
            if (regInfo == null)
            {
                return IntPtr.Zero; // Converted to #NUM
            }

            var ctx = XlDirectMarshal.GetMarshalContext();
            return ctx.ObjectReturn(regInfo);
        }

        internal static void CalculationCanceled()
        {
            if (_initialized)
                ExcelIntegration.CalculationCanceled();
        }

        internal static void CalculationEnded()
        {
            if (_initialized)
                ExcelIntegration.CalculationEnded();
        }
        #endregion
    }
}


