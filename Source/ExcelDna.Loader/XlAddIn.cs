//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Loader.Logging;

namespace ExcelDna.Loader
{
    using HRESULT = System.Int32;
    using IID = System.Guid;
    using CLSID = System.Guid;
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
        internal Int32 ExportInfoVersion; // Must be 8 for this version
        internal Int32 AppDomainId; // Id of the Sandbox AppDomain where the add-in runs.
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoOpen;
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoClose;
        internal IntPtr /* PFN_SHORT_VOID */            pXlAutoRemove;
        internal IntPtr /* PFN_VOID_LPXLOPER */         pXlAutoFree;
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
        internal Int32 ThunkTableLength;  // Must be EXPORT_COUNT
        internal IntPtr /*PFN*/ ThunkTable; // Actually (PFN ThunkTable[EXPORT_COUNT])
        #pragma warning restore 0649
    };

    // CAUTION: This type is loaded by reflection from the unmanaged loader.
    public unsafe static class XlAddIn
    {
        // This version must match the version declared in ExcelDna.Integration.ExcelIntegration
        const int ExcelIntegrationVersion = 8;

        static int thunkTableLength;
        static IntPtr thunkTable;

        // Passed in from unmanaged code during initialization 
        internal static IntPtr hModuleXll;

        static string pathXll;
        internal static string PathXll { get { return pathXll; } }

        static int xlCallVersion;
        internal static int XlCallVersion { get { return xlCallVersion; } }

        static bool _initialized = false;
        static bool _opened = false;

        #region Initialization

        public static bool Initialize32(int xlAddInExportInfoAddress, int hModuleXll, string pathXll)
        {
            return Initialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll);
        }

        public static bool Initialize64(long xlAddInExportInfoAddress, long hModuleXll, string pathXll)
        {
            return Initialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll);
        }

		private static unsafe bool Initialize(IntPtr xlAddInExportInfoAddress, IntPtr hModuleXll, string pathXll)
        {
            // NOTE: Too early for logging - the TraceSource in ExcelDna.Integration has not been initialized yet.
            Debug.Print("In sandbox AppDomain with Id: {0}, running on thread: {1}", AppDomain.CurrentDomain.Id, Thread.CurrentThread.ManagedThreadId);
            Debug.Assert(xlAddInExportInfoAddress != IntPtr.Zero, "InitializationInfo address is null");
            Debug.Print("InitializationInfo address: 0x{0:x8}", xlAddInExportInfoAddress);
			
			XlAddInExportInfo* pXlAddInExportInfo = (XlAddInExportInfo*)xlAddInExportInfoAddress;
            int exportInfoVersion = pXlAddInExportInfo->ExportInfoVersion;
            if (exportInfoVersion != 8)
            {
                Debug.Print("ExportInfoVersion '{0}' not supported", exportInfoVersion);
                return false;
            }

            fn_short_void fnXlAutoOpen = (fn_short_void)XlAutoOpen;
            GCHandle.Alloc(fnXlAutoOpen);
            pXlAddInExportInfo->pXlAutoOpen = Marshal.GetFunctionPointerForDelegate(fnXlAutoOpen);

            fn_short_void fnXlAutoClose = (fn_short_void)XlAutoClose;
            GCHandle.Alloc(fnXlAutoClose);
            pXlAddInExportInfo->pXlAutoClose = Marshal.GetFunctionPointerForDelegate(fnXlAutoClose);

            fn_short_void fnXlAutoRemove = (fn_short_void)XlAutoRemove;
            GCHandle.Alloc(fnXlAutoRemove);
            pXlAddInExportInfo->pXlAutoRemove = Marshal.GetFunctionPointerForDelegate(fnXlAutoRemove);

            fn_void_intptr fnXlAutoFree = (fn_void_intptr)XlAutoFree;
            GCHandle.Alloc(fnXlAutoFree);
            pXlAddInExportInfo->pXlAutoFree = Marshal.GetFunctionPointerForDelegate(fnXlAutoFree);

            fn_void_intptr fnXlAutoFree12 = (fn_void_intptr)XlAutoFree12;
            GCHandle.Alloc(fnXlAutoFree12);
            pXlAddInExportInfo->pXlAutoFree12 = Marshal.GetFunctionPointerForDelegate(fnXlAutoFree12);

            fn_void_intptr fnSetExcel12EntryPt = (fn_void_intptr)XlCallImpl.SetExcel12EntryPt;
            GCHandle.Alloc(fnSetExcel12EntryPt);
            pXlAddInExportInfo->pSetExcel12EntryPt = Marshal.GetFunctionPointerForDelegate(fnSetExcel12EntryPt);

            fn_hresult_void fnDllRegisterServer = (fn_hresult_void)DllRegisterServer;
            GCHandle.Alloc(fnDllRegisterServer);
            pXlAddInExportInfo->pDllRegisterServer = Marshal.GetFunctionPointerForDelegate(fnDllRegisterServer);

            fn_hresult_void fnDllUnregisterServer = (fn_hresult_void)DllUnregisterServer;
            GCHandle.Alloc(fnDllUnregisterServer);
            pXlAddInExportInfo->pDllUnregisterServer = Marshal.GetFunctionPointerForDelegate(fnDllUnregisterServer);

            fn_get_class_object fnDllGetClassObject = (fn_get_class_object)DllGetClassObject;
            GCHandle.Alloc(fnDllGetClassObject);
            pXlAddInExportInfo->pDllGetClassObject = Marshal.GetFunctionPointerForDelegate(fnDllGetClassObject);

            fn_hresult_void fnDllCanUnloadNow = (fn_hresult_void)DllCanUnloadNow;
            GCHandle.Alloc(fnDllCanUnloadNow);
            pXlAddInExportInfo->pDllCanUnloadNow = Marshal.GetFunctionPointerForDelegate(fnDllCanUnloadNow);

            fn_void_double fnSyncMacro = (fn_void_double)SyncMacro;
            GCHandle.Alloc(fnSyncMacro);
            pXlAddInExportInfo->pSyncMacro = Marshal.GetFunctionPointerForDelegate(fnSyncMacro);

            fn_intptr_intptr fnRegistrationInfo = (fn_intptr_intptr)RegistrationInfo;
            GCHandle.Alloc(fnRegistrationInfo);
            pXlAddInExportInfo->pRegistrationInfo = Marshal.GetFunctionPointerForDelegate(fnRegistrationInfo);

            fn_void_void fnCalculationCanceled = (fn_void_void)CalculationCanceled;
            GCHandle.Alloc(fnCalculationCanceled);
            pXlAddInExportInfo->pCalculationCanceled = Marshal.GetFunctionPointerForDelegate(fnCalculationCanceled);

            fn_void_void fnCalculationEnded = (fn_void_void)CalculationEnded;
            GCHandle.Alloc(fnCalculationEnded);
            pXlAddInExportInfo->pCalculationEnded = Marshal.GetFunctionPointerForDelegate(fnCalculationEnded);

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
			XlAddIn.hModuleXll = hModuleXll;
            XlAddIn.pathXll = pathXll;

            AssemblyManager.Initialize(hModuleXll, pathXll);

            try
            {
                LoadIntegration();
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

        private static void LoadIntegration()
        {
            // Get the assembly and ExcelIntegration type - will be loaded from file or from packed resources via AssemblyManager.AssemblyResolve.
            Assembly integrationAssembly = Assembly.Load("ExcelDna.Integration");
            Type integrationType = integrationAssembly.GetType("ExcelDna.Integration.ExcelIntegration");

            // Check the version declared in the ExcelIntegration class
            int integrationVersion = (int)integrationType.InvokeMember("GetExcelIntegrationVersion", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
            if (integrationVersion != ExcelIntegrationVersion)
            {
                // This is not the version we are expecting!
                throw new InvalidOperationException("Invalid ExcelIntegration version detected.");
            }

            // Get the methods that need to be called from the integration assembly
            MethodInfo tryExcelImplMethod = typeof(XlCallImpl).GetMethod("TryExcelImpl", BindingFlags.Static | BindingFlags.Public);
            Type tryExcelImplDelegateType = integrationAssembly.GetType("ExcelDna.Integration.TryExcelImplDelegate");
            Delegate tryExcelImplDelegate = Delegate.CreateDelegate(tryExcelImplDelegateType, tryExcelImplMethod);
            integrationType.InvokeMember("SetTryExcelImpl", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { tryExcelImplDelegate });

            MethodInfo registerMethodsMethod = typeof(XlRegistration).GetMethod("RegisterMethods", BindingFlags.Static | BindingFlags.Public);
            Type registerMethodsDelegateType = integrationAssembly.GetType("ExcelDna.Integration.RegisterMethodsDelegate");
            Delegate registerMethodsDelegate = Delegate.CreateDelegate(registerMethodsDelegateType, registerMethodsMethod);
            integrationType.InvokeMember("SetRegisterMethods", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { registerMethodsDelegate });

            MethodInfo registerWithAttMethod = typeof(XlRegistration).GetMethod("RegisterMethodsWithAttributes", BindingFlags.Static | BindingFlags.Public);
            Type registerWithAttDelegateType = integrationAssembly.GetType("ExcelDna.Integration.RegisterMethodsWithAttributesDelegate");
            Delegate registerWithAttDelegate = Delegate.CreateDelegate(registerWithAttDelegateType, registerWithAttMethod);
            integrationType.InvokeMember("SetRegisterMethodsWithAttributes", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { registerWithAttDelegate });

            MethodInfo registerDelAttMethod = typeof(XlRegistration).GetMethod("RegisterDelegatesWithAttributes", BindingFlags.Static | BindingFlags.Public);
            Type registerDelAttDelegateType = integrationAssembly.GetType("ExcelDna.Integration.RegisterDelegatesWithAttributesDelegate");
            Delegate registerDelAttDelegate = Delegate.CreateDelegate(registerDelAttDelegateType, registerDelAttMethod);
            integrationType.InvokeMember("SetRegisterDelegatesWithAttributes", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { registerDelAttDelegate });

            MethodInfo getResourceBytesMethod = typeof(AssemblyManager).GetMethod("GetResourceBytes", BindingFlags.Static | BindingFlags.NonPublic);
            Type getResourceBytesDelegateType = integrationAssembly.GetType("ExcelDna.Integration.GetResourceBytesDelegate");
            Delegate getResourceBytesDelegate = Delegate.CreateDelegate(getResourceBytesDelegateType, getResourceBytesMethod);
            integrationType.InvokeMember("SetGetResourceBytesDelegate", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { getResourceBytesDelegate });

            // set up helpers for future calls
            IntegrationHelpers.Bind(integrationAssembly, integrationType);
            IntegrationMarshalHelpers.Bind(integrationAssembly);
        }
        
        private static void InitializeIntegration()
        {
            if (!_initialized)
            {
                IntegrationHelpers.InitializeIntegration(pathXll);
                TraceLogger.IntegrationTraceSource = IntegrationHelpers.GetIntegrationTraceSource();
                _initialized = true;
            }
        }

        private static void DeInitializeIntegration()
        {
            if (_initialized)
            {
                if (_opened)
                {
                    IntegrationHelpers.DnaLibraryAutoClose();
                    XlRegistration.UnregisterMethods();
                }
                TraceLogger.IntegrationTraceSource = null;
                IntegrationHelpers.DeInitializeIntegration();
                _initialized = false;
                _opened = false;
            }
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
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignore*/ , true, "Registering library " + pathXll);
				InitializeIntegration();
                Logger.Initialization.Verbose("In XlAddIn.XlAutoOpen");
                
                // v. 30 - moved the setting of _opened before calling AutoOpen, 
                // so that checking in DeInitializeIntegration does not prevent AutoOpen - unloading via xlAutoRemove from working.
                _opened = true;

                // InitializeIntegration has loaded the DnaLibrary
                IntegrationHelpers.DnaLibraryAutoOpen();

                result = 1; // All is OK
            }
            catch (Exception e)
            {
                // Can't use logging here
                string alertMessage = string.Format("A problem occurred while an add-in was being initialized (InitializeIntegration failed - {1}).\r\nThe add-in is built with ExcelDna and is being loaded from {0}", pathXll, e.Message);
				object xlCallResult;
				XlCallImpl.TryExcelImpl(XlCallImpl.xlcAlert, out xlCallResult /*Ignored*/, alertMessage , 3 /* Only OK Button, Warning Icon*/);
                result = 0;
            }
            finally
            {
                // Clear the status bar message
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignored*/ , false);
                // Debug.Print("Clear status bar message result: " + xlCallResult);
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

		internal static void XlAutoFree(IntPtr pXloper)
        {
            // CONSIDER: This might be improved....
            // Another option would be to have the Com memory allocator run in unmanaged code.
            // Right now I think this is OK, and easiest from where I'm coming.
            // This function can only be called after a return from a user function.
            // I just free all the possibly big memory allocations.

            XlObjectArrayMarshaler.FreeMemory();
        }

		internal static void XlAutoFree12(IntPtr pXloper12)
        {
            // CONSIDER: This might be improved....
            // Another option would be to have the Com memory allocator run in unmanaged code.
            // Right now I think this is OK, and easiest from where I'm coming.
            // This function can only be called after a return from a user function.
            // I just free all the possibly big memory allocations.

            XlObjectArray12Marshaler.FreeMemory();
        }

        // Note: XlAddInManagerInfo is now implemented in the unmanaged side (for performance in the add-in dialog).
        //internal static IntPtr XlAddInManagerInfo(IntPtr pXloperAction)
        //{
        //    Debug.WriteLine("In XlAddIn.XlAddInManagerInfo");
        //    ICustomMarshaler m = XlObjectMarshaler.GetInstance("");
        //    object action = m.MarshalNativeToManaged(pXloperAction);
        //    object result;
        //    if ((action is double && (double)action == 1.0))
        //    {
        //        InitializeIntegration();
        //        result = IntegrationHelpers.DnaLibraryGetName();
        //    }
        //    else
        //        result = IntegrationMarshalHelpers.GetExcelErrorObject(IntegrationMarshalHelpers.ExcelError_ExcelErrorValue);
        //    return m.MarshalManagedToNative(result);
        //}

        //internal static IntPtr XlAddInManagerInfo12(IntPtr pXloperAction12)
        //{
        //    Debug.WriteLine("In XlAddIn.XlAddInManagerInfo12");
        //    ICustomMarshaler m = XlObject12Marshaler.GetInstance("");
        //    object action = m.MarshalNativeToManaged(pXloperAction12);
        //    object result;
        //    if ((action is double && (double)action == 1.0))
        //    {
        //        InitializeIntegration();
        //        result = IntegrationHelpers.DnaLibraryGetName();
        //    }
        //    else
        //        result = IntegrationMarshalHelpers.GetExcelErrorObject(IntegrationMarshalHelpers.ExcelError_ExcelErrorValue);
        //    return m.MarshalManagedToNative(result);
        //}

        #endregion

        #region Com Server exports
        internal static HRESULT DllRegisterServer()
        {
            InitializeIntegration();
            return IntegrationHelpers.DllRegisterServer();
        }

        internal static HRESULT DllUnregisterServer()
        {
            InitializeIntegration();
            return IntegrationHelpers.DllUnregisterServer();
        }

        internal static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        {
            Debug.Print("DllGetClassObject entered - calling InitializeIntegration.");
            HRESULT result;
            InitializeIntegration();
            Logger.Initialization.Verbose("In DllGetClassObject");
            result = IntegrationHelpers.DllGetClassObject(clsid, iid, out ppunk);
            return result;
        }

        internal static HRESULT DllCanUnloadNow()
        {
            InitializeIntegration();
            return IntegrationHelpers.DllCanUnloadNow();
        }
        #endregion

        #region Extensions support
        internal static void SyncMacro(double dValue)
        {
            if (_initialized)
                IntegrationHelpers.SyncMacro(dValue);
        }

        internal static IntPtr RegistrationInfo(IntPtr pParam)
        {
            if (!_initialized)
            {
                return IntPtr.Zero;
            }

            // CONSIDER: This might not be the right place for this.
            ICustomMarshaler m = XlObject12Marshaler.GetInstance("");
            object param = m.MarshalNativeToManaged(pParam);
            object regInfo = XlRegistration.GetRegistrationInfo(param);
            if (regInfo == null)
            {
                return IntPtr.Zero; // Converted to #NUM
            }

            return m.MarshalManagedToNative(regInfo);
        }

        internal static void CalculationCanceled()
        {
            if (_initialized)
                IntegrationHelpers.CalculationCanceled();
        }

        internal static void CalculationEnded()
        {
            if (_initialized)
                IntegrationHelpers.CalculationEnded();
        }
        #endregion
    }
}


