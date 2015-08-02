//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

namespace ExcelDna.Integration
{
    using ExcelDna.ComInterop;
    using ExcelDna.Logging;
    using HRESULT = Int32;

    // CAUTION: These functions are called _via reflection_ by
    // ExcelDna.Loader.XlLibrary to set up the link between the loader 
    // and the integration library.
    // Signatures, private/public etc. is fragile.

    internal delegate int TryExcelImplDelegate(int xlFunction, out object result, params object[] parameters);
    internal delegate void RegisterMethodsDelegate(List<MethodInfo> methods);
    internal delegate void RegisterMethodsWithAttributesDelegate(List<MethodInfo> methods, List<object> functionAttributes, List<List<object>> argumentAttributes);
    internal delegate void RegisterDelegatesWithAttributesDelegate(List<Delegate> delegates, List<object> functionAttributes, List<List<object>> argumentAttributes);
    internal delegate byte[] GetResourceBytesDelegate(string resourceName, int type); // types: 0 - Assembly, 1 - Dna file, 2 - Image
    internal delegate void SyncMacroDelegate(double dValue);
	public delegate object UnhandledExceptionHandler(object exceptionObject);

    public static class ExcelIntegration
    {
        // This version must match the version declared in ExcelDna.Loader.XlAddIn.
        const int ExcelIntegrationVersion = 8;

        private static TryExcelImplDelegate tryExcelImpl;
        internal static void SetTryExcelImpl(TryExcelImplDelegate d)
        {
            tryExcelImpl = d;
        }

        internal static XlCall.XlReturn TryExcelImpl(int xlFunction, out object result, params object[] parameters)
        {
            if (tryExcelImpl != null)
            {
                return (XlCall.XlReturn)tryExcelImpl(xlFunction, out result, parameters);
            }
            result = null;
            return XlCall.XlReturn.XlReturnFailed;
        }

        private static RegisterMethodsDelegate registerMethods;
        internal static void SetRegisterMethods(RegisterMethodsDelegate d)
        {
            registerMethods = d;
        }

        private static RegisterMethodsWithAttributesDelegate registerMethodsWithAttributes;
        internal static void SetRegisterMethodsWithAttributes(RegisterMethodsWithAttributesDelegate d)
        {
            registerMethodsWithAttributes = d;
        }

        private static RegisterDelegatesWithAttributesDelegate registerDelegatesWithAttributes;
        internal static void SetRegisterDelegatesWithAttributes(RegisterDelegatesWithAttributesDelegate d)
        {
            registerDelegatesWithAttributes = d;
        }

        // These are the public 'externally' exposed members.

        // Get the assemblies that were considered for registration - both ExternalLibraries and Projects or code from the .dna file.
        // This is not used internally, but available for custom registration.
        public static IEnumerable<Assembly> GetExportedAssemblies()
        {
            return DnaLibrary.CurrentLibrary.GetExportedAssemblies();
        }

        public static void RegisterMethods(List<MethodInfo> methods)
        {
            registerMethods(methods);
        }

        public static void RegisterMethods(List<MethodInfo> methods,
                                           List<object> methodAttributes,
                                           List<List<object>> argumentAttributes)
        {
            ClearExplicitRegistration(methodAttributes);
            registerMethodsWithAttributes(methods, methodAttributes, argumentAttributes);
        }

        public static void RegisterDelegates(List<Delegate> delegates,
                                             List<object> methodAttributes,
                                             List<List<object>> argumentAttributes)
        {
            ClearExplicitRegistration(methodAttributes);
            registerDelegatesWithAttributes(delegates, methodAttributes, argumentAttributes);
        }

        // Fix up the ExplicitRegistration, since we _are_ now explicitly registering
        static void ClearExplicitRegistration(List<object> methodAttributes)
        {
            foreach (object attrib in methodAttributes)
            {
                ExcelFunctionAttribute funcAttrib = attrib as ExcelFunctionAttribute;
                if (funcAttrib != null)
                {
                    funcAttrib.ExplicitRegistration = false;
                    continue;
                }
                ExcelCommandAttribute cmdAttrib = attrib as ExcelCommandAttribute;
                if (cmdAttrib != null)
                {
                    cmdAttrib.ExplicitRegistration = false;
                }
            }
        }

		private static UnhandledExceptionHandler unhandledExceptionHandler;
		public static void RegisterUnhandledExceptionHandler(UnhandledExceptionHandler h)
		{
			unhandledExceptionHandler = h;
		}


        #region Registration Info
        // Public function to get registration info for this or other Excel-DNA .xlls
        // Return null if the matching RegistrationInfo function is not found.
        public static object GetRegistrationInfo(string xllPath, double registrationUpdateVersion)
        {
            return RegistrationInfo.GetRegistrationInfo(xllPath, registrationUpdateVersion);
        }

        // Just added for symmetry
        /// <summary>
        /// 
        /// </summary>
        /// <param name="xllPath"></param>
        /// <returns>Either a string with the name of the XLL, or the ExcelError.ExcelErrorValue error.</returns>
        public static object RegisterXLL(string xllPath)
        {
            Debug.Print("## Registering Add-In: " + xllPath);
            return XlCall.Excel(XlCall.xlfRegister, xllPath);
        }

        public static void UnregisterXLL(string xllPath)
        {
            Debug.Print("## Unregistering Add-In: " + xllPath);
            // Little detour to get Excel-DNA to fully unregister the function names.
            object removeId = XlCall.Excel(XlCall.xlfRegister, xllPath, "xlAutoRemove", "I", ExcelEmpty.Value, ExcelEmpty.Value, 2);
            object removeResult = XlCall.Excel(XlCall.xlfCall, removeId);
            object removeUnregisterResult = XlCall.Excel(XlCall.xlfUnregister, removeId);
            XlCall.Excel(XlCall.xlfUnregister, xllPath);
        }
        #endregion


		// WARNING: This method is bound by name from the ExcelDna.Loader in IntegrationHelpers.Bind.
		// It should not throw an exception, and is called directly from the UDF exceptionhandler.
		internal static object HandleUnhandledException(object exceptionObject)
		{
			if (unhandledExceptionHandler == null)
			{
				return ExcelError.ExcelErrorValue;
			}
			try
			{
				return unhandledExceptionHandler(exceptionObject);
			}
			catch (Exception ex)
			{
				Debug.WriteLine("Exception in UnhandledExceptionHandler: " + ex);
				return ExcelError.ExcelErrorValue;
			}
		}

        private static GetResourceBytesDelegate getResourceBytesDelegate;
        internal static void SetGetResourceBytesDelegate(GetResourceBytesDelegate d)
        {
            getResourceBytesDelegate = d;
        }

		internal static byte[] GetAssemblyBytes(string assemblyName)
		{
			return getResourceBytesDelegate(assemblyName, 0);
		}

		internal static byte[] GetDnaFileBytes(string dnaFileName)
		{
			return getResourceBytesDelegate(dnaFileName, 1);
		}

        internal static byte[] GetImageBytes(string imageName)
        {
            return getResourceBytesDelegate(imageName, 2);
        }

        internal static byte[] GetSourceBytes(string sourceName)
        {
            return getResourceBytesDelegate(sourceName, 3);
        }

        // Called via Reflection from Loader
        internal static void Initialize(string xllPath)
        {
			ExcelDnaUtil.Initialize();  // Set up window handle
            Logging.TraceLogger.Initialize();
            DnaLibrary.InitializeRootLibrary(xllPath);
        }

        // Called via Reflection from Loader
        internal static void DeInitialize()
        {
            DnaLibrary.DeInitialize();
        }

        internal static void DnaLibraryAutoOpen()
        {
			Logger.Initialization.Verbose("Enter Integration.DnaLibraryAutoOpen");
			try
			{
				DnaLibrary.CurrentLibrary.AutoOpen();
            }
			catch (Exception e)
			{
                Logger.Initialization.Error(e, "Integration.DnaLibraryAutoOpen Error");
			}
            Logger.Initialization.Verbose("Exit Integration.DnaLibraryAutoOpen");
		}

        internal static void DnaLibraryAutoClose()
        {
            DnaLibrary.CurrentLibrary.AutoClose();
        }

        internal static string DnaLibraryGetName()
        {
            return DnaLibrary.CurrentLibrary.Name;
        }

        // ComServer related exports just delegates to ComServer class.
        internal static HRESULT DllRegisterServer()
        {
            try
            {
                return ComServer.DllRegisterServer();
            }
            catch (UnauthorizedAccessException uae)
            {
                Debug.Write("DllRegisterServer error: " + uae.Message);
                // Expected only if we can't write to HKCU\Software\Classes.
                return ComAPI.E_ACCESSDENIED;
            }
        }

        internal static HRESULT DllUnregisterServer()
        {
            try
            {
                return ComServer.DllUnregisterServer();
            }
            catch (UnauthorizedAccessException uae)
            {
                Debug.Write("DllRegisterServer error: " + uae.Message);
                // Expected only if we can't write to HKCU\Software\Classes.
                return ComAPI.E_ACCESSDENIED;
            }
        }

        // internal static HRESULT DllGetClassObject([In] ref CLSID rclsid, [In] ref IID riid, [Out, MarshalAs(UnmanagedType.Interface)] out object ppunk)
        internal static HRESULT DllGetClassObject(Guid clsid, Guid iid, out IntPtr ppunk)
        {
            return ComServer.DllGetClassObject(clsid, iid, out ppunk);
        }

        internal static HRESULT DllCanUnloadNow()
        {
            return ComServer.DllCanUnloadNow();
        }

        // Implementation for SyncMacro
        // CONSIDER: This could be a more direct registration?
        static SyncMacroDelegate syncMacro = null;
        internal static void SetSyncMacro(SyncMacroDelegate d)
        {
            syncMacro = d;
        }

        // Called via Reflection from Loader
        internal static void SyncMacro(double dValue)
        {
            if (syncMacro != null)
                syncMacro(dValue);
        }

        // Called via Reflection from Loader
        internal static void CalculationCanceled()
        {
            ExcelAsyncUtil.OnCalculationCanceled();    
        }

        // Called via Reflection from Loader
        internal static void CalculationEnded()
        {
            ExcelAsyncUtil.OnCalculationEnded();    
        }

        // Called via Reflection from Loader after Initialization
        internal static TraceSource GetIntegrationTraceSource()
        {
            return Logging.TraceLogger.IntegrationTraceSource;
        }

        // This version check is made by the ExceDna.Loader to make sure we have matching versions.
        internal static int GetExcelIntegrationVersion()
        {
            return ExcelIntegrationVersion;
        }
    }

    #region Obsolete classes
    [Obsolete("Use ExcelDna.Integration.ExcelIntegration class")]
    public class XlLibrary
    {
        [Obsolete("Use ExcelDna.Integration.Integration.RegisterMethods method")]
        public static void RegisterMethods(List<MethodInfo> methods)
        {
            ExcelIntegration.RegisterMethods(methods);
        }
    }

    [Obsolete("Use class ExcelDna.Integration.ExcelIntegration instead.")]
    public static class Integration
    {
        public static void RegisterMethods(List<MethodInfo> methods)
        {
            ExcelIntegration.RegisterMethods(methods);
        }

        public static void RegisterUnhandledExceptionHandler(UnhandledExceptionHandler h)
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(h);
        }
    }
    #endregion
}
