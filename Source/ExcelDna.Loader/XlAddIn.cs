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
using System.IO;
using System.Text;
using System.Reflection;
using System.Runtime.Remoting;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Threading;

namespace ExcelDna.Loader
{
    using HRESULT = System.Int32;
    using IID = System.Guid;
    using CLSID = System.Guid;
    internal delegate short fn_short_void();
    internal delegate void fn_void_intptr(IntPtr intPtr);
    internal delegate IntPtr fn_intptr_intptr(IntPtr intPtr);
    internal delegate HRESULT fn_hresult_void();
    internal delegate HRESULT fn_get_class_object(CLSID rclsid, IID riid, out IntPtr ppunk);

    // CAUTION: This struct is also defined in the unmanaged loader.
    internal struct XlAddInExportInfo
    {
        #pragma warning disable 0649 // Field 'field' is never assigned to, and will always have its default value 'value'
        internal Int32 ExportInfoVersion; // Must be 4 for this version
        internal Int32 AppDomainId; // Id of the Sandbox AppDomain where the add-in runs.
        internal IntPtr /* PFN_SHORT_VOID */ pXlAutoOpen;
        internal IntPtr /* PFN_SHORT_VOID */ pXlAutoClose;
        internal IntPtr /* PFN_SHORT_VOID */ pXlAutoRemove;
        internal IntPtr /* PFN_VOID_LPXLOPER */     pXlAutoFree;
        internal IntPtr /* PFN_VOID_LPXLOPER12 */   pXlAutoFree12;
        internal IntPtr /* PFN_PFNEXCEL12 */ pSetExcel12EntryPt;
        internal IntPtr /* PFN_HRESULT_VOID */ pDllRegisterServer;
        internal IntPtr /* PFN_HRESULT_VOID */ pDllUnregisterServer;
        internal IntPtr /* PFN_GET_CLASS_OBJECT */ pDllGetClassObject;
        internal IntPtr /* PFN_HRESULT_VOID */ pDllCanUnloadNow;
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
        static string pathXll;
        internal static IntPtr hModuleXll;

        static int xlCallVersion;
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
         
            Debug.Print("Initialize - in sandbox AppDomain with Id: {0}, running on thread: {1}", AppDomain.CurrentDomain.Id, Thread.CurrentThread.ManagedThreadId);
            Debug.Assert(xlAddInExportInfoAddress != IntPtr.Zero);
            Debug.Print("InitializationInfo Address: 0x{0:x8}", xlAddInExportInfoAddress);
			
			XlAddInExportInfo* pXlAddInExportInfo = (XlAddInExportInfo*)xlAddInExportInfoAddress;
            if (pXlAddInExportInfo->ExportInfoVersion != 5)
            {
                Debug.Print("ExportInfoVersion not supported.");
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

            // Thunk table for registered functions
            thunkTableLength = pXlAddInExportInfo->ThunkTableLength;
            thunkTable = pXlAddInExportInfo->ThunkTable;

			// This is the place where we call into Excel - this causes SecurityPermission exception
			// when run from VSTO. I don't know how to deal with this problem yet.
			// TODO: Learn more about the special security stuff in VSTO.
			try
			{
				XlAddIn.xlCallVersion = XlCallImpl.XLCallVer() / 256;
			}
			catch (Exception e)
			{
				Debug.WriteLine("XlAddIn: XLCallVer Exception: " + e);

                // CONSIDER: Is this right / needed ?
                // As a test for the HPC support, I ignore error here and just assume we are under new Excel.
                // This will cause the common error here to get pushed to later ...
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
            catch (Exception e)
            {
                Debug.WriteLine("XlAddIn: Initialize Exception: " + e);
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
            Assembly integrationAssembly = Assembly.Load("ExcelDna.Integration");
            Type integrationType = integrationAssembly.GetType("ExcelDna.Integration.Integration");

            // Get the methods that need to be called from the integration assembly
            MethodInfo tryExcelImplMethod = typeof(XlCallImpl).GetMethod("TryExcelImpl", BindingFlags.Static | BindingFlags.Public);
            Type tryExcelImplDelegateType = integrationAssembly.GetType("ExcelDna.Integration.TryExcelImplDelegate");
            Delegate tryExcelImplDelegate = Delegate.CreateDelegate(tryExcelImplDelegateType, tryExcelImplMethod);
            integrationType.InvokeMember("SetTryExcelImpl", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { tryExcelImplDelegate });

            MethodInfo registerMethodsMethod = typeof(XlAddIn).GetMethod("RegisterMethods", BindingFlags.Static | BindingFlags.Public);
            Type registerMethodsDelegateType = integrationAssembly.GetType("ExcelDna.Integration.RegisterMethodsDelegate");
            Delegate registerMethodsDelegate = Delegate.CreateDelegate(registerMethodsDelegateType, registerMethodsMethod);
            integrationType.InvokeMember("SetRegisterMethods", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { registerMethodsDelegate });

            MethodInfo getResourceBytesMethod = typeof(AssemblyManager).GetMethod("GetResourceBytes", BindingFlags.Static | BindingFlags.NonPublic);
            Type getResourceBytesDelegateType = integrationAssembly.GetType("ExcelDna.Integration.GetResourceBytesDelegate");
            Delegate getResourceBytesDelegate = Delegate.CreateDelegate(getResourceBytesDelegateType, getResourceBytesMethod);
            integrationType.InvokeMember("SetGetResourceBytesDelegate", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] { getResourceBytesDelegate });

            // set up helpers for future calls
            IntegrationHelpers.Bind(integrationAssembly);
            IntegrationMarshalHelpers.Bind(integrationAssembly);
        }

        private static void InitializeIntegration()
        {
            if (!_initialized)
            {
                IntegrationHelpers.InitializeIntegration(pathXll);
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
                    UnregisterMethods();
                }
                IntegrationHelpers.DeInitializeIntegration();
                _initialized = false;
                _opened = false;
            }
        }
        #endregion

        #region Managed Xlxxxx functions
        internal static short XlAutoOpen()
        {
            Debug.Print("AppDomain Id: " + AppDomain.CurrentDomain.Id + " (Default: " + AppDomain.CurrentDomain.IsDefaultAppDomain() + ")");
			short result = 0;
            try
            {
                Debug.WriteLine("In XlAddIn.XlAutoOpen");
                if (_opened)
                {
                    DeInitializeIntegration();
                }
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignore*/ , true, "Registering library " + pathXll);
				InitializeIntegration();
                // InitializeIntegration has loaded the DnaLibrary
                IntegrationHelpers.DnaLibraryAutoOpen();
                _opened = true;
                result = 1; // All is OK
            }
            catch (Exception e)
            {
                // TODO: What to do here - maybe prefer Trace...?
                Debug.WriteLine("ExcelDna.Loader.XlAddin.XlAutoOpen. Exception during Integration load: " + e.ToString());
				string alertMessage = string.Format("A problem occurred while an add-in was being initialized (InitializeIntegration failed).\r\nThe add-in is built with ExcelDna and is being loaded from {0}", pathXll);
				object xlCallResult;
				XlCallImpl.TryExcelImpl(XlCallImpl.xlcAlert, out xlCallResult /*Ignored*/, alertMessage , 3 /* Only OK Button, Warning Icon*/);
                result = 0;
            }
            finally
            {
                // Clear the status bar message
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlcMessage, out xlCallResult /*Ignored*/ , false);
            }
            return result;
        }

		internal static short XlAutoClose()
        {
            short result = 0;
            try
            {
                Debug.WriteLine("In XlAddIn.XlAutoClose");
                // This also gets called when workbook starts closing - and can still be cancelled
                result = 1; // 0 if problems ?
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
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
                Debug.WriteLine("In XlAddIn.XlAutoAdd");
                result = 1;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }

            return result;
        }

		internal static short XlAutoRemove()
        {
            short result = 0;
            try
            {
                Debug.WriteLine("In XlAddIn.XlAutoRemove");
                // Apparently better if called here, 
                // so I try to, but make it safe to call again.
                DeInitializeIntegration();
                return 1; // 0 if problems ?
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
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
            Debug.Print("DllGetClassObject reached!!");
           // IntPtr pout;
            HRESULT result;
            InitializeIntegration();
            result = IntegrationHelpers.DllGetClassObject(clsid, iid, out ppunk);
            return result;
        }

        internal static HRESULT DllCanUnloadNow()
        {
            InitializeIntegration();
            return IntegrationHelpers.DllCanUnloadNow();
        }
        #endregion

        // TODO: Migrate registration to ExcelDna.Integration
        #region Managed Registration
        static List<XlMethodInfo> registeredMethods = new List<XlMethodInfo>();
        static List<string> addedMenus = new List<string>();
        static List<XlMethodInfo> addedCommands = new List<XlMethodInfo>();

        public static void RegisterMethods(List<MethodInfo> methods)
        {
            List<XlMethodInfo> xlMethods = XlMethodInfo.ConvertToXlMethodInfos(methods);
            xlMethods.ForEach(RegisterXlMethod);
        }

        private static void RegisterXlMethod(XlMethodInfo mi)
        {
            int index = registeredMethods.Count;
            SetJump(index, mi.FunctionPointer);
            String procName = String.Format("f{0}", index);

            string functionType;
            if ( mi.IsCommand)
            {
                if (mi.Parameters.Length == 0)
                    functionType = "";  // OK since no other types will be added
                else
                    functionType = ">"; // Use the void / inplace indicator if needed.
            }
            else
                functionType = mi.ReturnType.XlType;

            string argumentNames = "";
            bool showDescriptions = false;
            string[] argumentDescriptions = new string[mi.Parameters.Length];
            string helpTopic;

            for (int j = 0; j < mi.Parameters.Length; j++)
            {
                XlParameterInfo pi = mi.Parameters[j];

                functionType += pi.XlType;
                if (j > 0)
                    argumentNames += ",";
                argumentNames += pi.Name;
                argumentDescriptions[j] = pi.Description;

                if (pi.Description != "")
                    showDescriptions = true;

                // DOCUMENT: Here is the patch for the Excel Function Description bug.
                // DOCUMENT: I add ". " to the last parameters.
                if (j == mi.Parameters.Length - 1)
                    argumentDescriptions[j] += ". ";

            } // for each parameter

            if (mi.IsClusterSafe && ProcessHelper.SupportsClusterSafe)
                functionType += "&"; 
            
            if (mi.IsMacroType)
                functionType += "#";

            if (!mi.IsMacroType && mi.IsThreadSafe && XlAddIn.xlCallVersion >= 12)
                functionType += "$";

            if (mi.IsVolatile)
                functionType += "!";
            // DOCUMENT: If # is set and there is an R argument, Excel considers the function volatile anyway.
            // You can call xlfVolatile, false in beginning of function to clear.

			// DOCUMENT: There is a bug? in Excel 2007 that limits the total argumentname string to 255 chars.
            // TODO: Check whether this is fixed in Excel 2010 yet.
			// DOCUMENT: I truncate the argument string for all versions.
			if (argumentNames.Length > 255)
				argumentNames = argumentNames.Substring(0, 255);

			// DOCUMENT: Here is the patch for the Excel Function Description bug.
            // DOCUMENT: I add ". " if the function takes no parameters and has a description.
            string functionDescription = mi.Description;
            if (mi.Parameters.Length == 0 && functionDescription != "")
                functionDescription += ". ";

			// DOCUMENT: When there is no description, we don't add any.
			// This allows the user to work around the Excel bug where an extra parameter is displayed if
			// the function has no parameter but displays a description
			if (mi.Description != "")
                showDescriptions = true;

			int numArguments;
            // DOCUMENT: Maximum 20 Argument Descriptions when registering using Excel4 function.
            int maxDescriptions = (XlAddIn.xlCallVersion < 12) ? 20 : 245;
            int numArgumentDescriptions;
            if (showDescriptions)
            {
                numArgumentDescriptions = Math.Min(argumentDescriptions.Length, maxDescriptions);
                numArguments = 10 + numArgumentDescriptions;
            }
            else
            {
                numArgumentDescriptions = 0;
                numArguments = 9;
            }

            // Make HelpTopic without full path relative to xllPath
            if (string.IsNullOrEmpty(mi.HelpTopic))
            {
                helpTopic = mi.HelpTopic;
            }
            else
            {
                // DOCUMENT: If HelpTopic is not rooted - it is expanded relative to .xll path.
                if (Path.IsPathRooted(mi.HelpTopic))
                {
                    helpTopic = mi.HelpTopic;
                }
                else
                {
                    helpTopic = Path.Combine(Path.GetDirectoryName(pathXll), mi.HelpTopic);
                }
            }

            object[] registerParameters = new object[numArguments];
            registerParameters[0] = pathXll;
            registerParameters[1] = procName;
            registerParameters[2] = functionType;
            registerParameters[3] = mi.Name;
            registerParameters[4] = argumentNames;
            registerParameters[5] = mi.IsCommand ? 2 /*macro*/
                                                          : (mi.IsHidden ? 0 : 1); /*function*/
            registerParameters[6] = mi.Category;
            registerParameters[7] = mi.ShortCut; /*shortcut_text*/
            registerParameters[8] = helpTopic; /*help_topic*/ ;

            if (showDescriptions)
            {
                registerParameters[9] = functionDescription;

                for (int k = 0; k < numArgumentDescriptions; k++)
                {
                    registerParameters[10 + k] = argumentDescriptions[k];
                }
            }

            // Basically suppress problems here !?
            try
            {
                object xlCallResult;
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfRegister, out xlCallResult, registerParameters);
                Debug.Print("Register - XllPath={0}, ProcName={1}, FunctionType={2}, MethodName={3} - Result={4}", registerParameters[0], registerParameters[1], registerParameters[2], registerParameters[3], xlCallResult);
                if (xlCallResult is double)
                {
                    mi.RegisterId = (double)xlCallResult;
                    registeredMethods.Add(mi);
                    if (mi.IsCommand)
                    {
                        RegisterMenu(mi);
                    }
                }
                else
                {
                    // TODO: What to do here? LogDisplay??
                    Debug.Print("Registration Error! - Register call failed for method {0}", mi.Name);
                }
            }
            catch (Exception e)
            {
                // TODO: What to do here? LogDisplay??
                Debug.WriteLine("Registration Error! - " + e.Message);
            }
        }

        private static void RegisterMenu(XlMethodInfo mi)
        {
            if (mi.MenuName != null && mi.MenuName != ""
                && mi.MenuText != null && mi.MenuText != "")
            {
                IntegrationHelpers.AddCommandMenu(mi.Name, mi.MenuName, mi.MenuText, mi.Description, mi.ShortCut, mi.HelpTopic);
            }
        }

        internal static void UnregisterMethods()
        {
            object xlCallResult;

            // Remove menus
            IntegrationHelpers.RemoveCommandMenus();

            // Now take out the methods
            foreach (XlMethodInfo mi in registeredMethods)
            {
                if (!mi.IsCommand)
                {
                    // I follow the advice from X-Cell website
                    // to get function out of Wizard
                    XlCallImpl.TryExcelImpl(XlCallImpl.xlfRegister, out xlCallResult, pathXll, "xlAutoRemove", "J", mi.Name, IntegrationMarshalHelpers.GetExcelMissingValue(), 0);
                }
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfSetName, out xlCallResult, mi.Name);
                XlCallImpl.TryExcelImpl(XlCallImpl.xlfUnregister, out xlCallResult, mi.RegisterId);
            }
            registeredMethods.Clear();
        }

        internal static int XlCallVersion
        {
            get { return xlCallVersion; }
        }

        #endregion

    }
    
    public static class AppDomainHelper
    {
        // This method is called from unmanaged code in a temporary AppDomain, just to be able to call
        // the right AppDomain.CreateDomain overload.
        public static AppDomain CreateFullTrustSandbox()
        {
            try
            {
                Debug.Print("CreateSandboxAndInitialize - in loader AppDomain with Id: " + AppDomain.CurrentDomain.Id);

                PermissionSet pset = new PermissionSet(PermissionState.Unrestricted);
                AppDomainSetup loaderAppDomainSetup = AppDomain.CurrentDomain.SetupInformation;
                AppDomainSetup sandboxAppDomainSetup = new AppDomainSetup();
                sandboxAppDomainSetup.ApplicationName = loaderAppDomainSetup.ApplicationName;
                sandboxAppDomainSetup.ConfigurationFile = loaderAppDomainSetup.ConfigurationFile;
                sandboxAppDomainSetup.ApplicationBase = loaderAppDomainSetup.ApplicationBase;
                sandboxAppDomainSetup.ShadowCopyFiles = loaderAppDomainSetup.ShadowCopyFiles;
                sandboxAppDomainSetup.ShadowCopyDirectories = loaderAppDomainSetup.ShadowCopyDirectories;

                // create the sandboxed domain
                AppDomain sandbox = AppDomain.CreateDomain(
                    "FullTrustSandbox(" + AppDomain.CurrentDomain.FriendlyName + ")",
                    null,
                    sandboxAppDomainSetup,
                    pset);

                Debug.Print("CreateFullTrustSandbox - sandbox AppDomain created. Id: " + sandbox.Id);

                return sandbox;
            }
            catch (Exception ex)
            {
                Debug.Print("Error during CreateFullTrustSandbox: " + ex.ToString());
                return AppDomain.CurrentDomain;
            }

        }
    }

    internal static class ProcessHelper
    {
        private static bool _isInitialized = false;
        private static bool _isRunningOnCluster;
        private static int _processMajorVersion;

        public static bool IsRunningOnCluster
        {
            get
            {
                EnsureInitialized();
                return _isRunningOnCluster;
            }
        }

        public static int ProcessMajorVersion
        {
            get
            {
                EnsureInitialized();
                return _processMajorVersion;
            }
        }

        public static bool SupportsClusterSafe
        {
            get
            {
                return IsRunningOnCluster || (ProcessMajorVersion >= 14);
            }
        }

        private static void EnsureInitialized()
        {
            if (!_isInitialized)
            {
                Process hostProcess = Process.GetCurrentProcess();
                _isRunningOnCluster = !(hostProcess.ProcessName.Equals("EXCEL", StringComparison.InvariantCultureIgnoreCase));
                _processMajorVersion = hostProcess.MainModule.FileVersionInfo.FileMajorPart;

                _isInitialized = true;
            }
        }            
    }
}


