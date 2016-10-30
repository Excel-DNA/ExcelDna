//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Reflection;
using System.Reflection.Emit;

namespace ExcelDna.Loader
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;

    // This class has a friend in XlCustomMarshal.
    // TODO: Nothing here is performance-critical, but the Invoke calls might be quite slow... use delegates?
    public static class IntegrationHelpers
    {
        static Type integrationType;
        static MethodInfo addCommandMenu;
        static MethodInfo removeCommandMenus;
        internal delegate object ExceptionHandler(object ex);
		internal static ExceptionHandler UnhandledExceptionHandler;
        
        public static object HandleUnhandledException(object exception)
        {
            return UnhandledExceptionHandler(exception);
        }

        internal static void Bind(Assembly integrationAssembly, Type bindIntegrationType)
        {
            integrationType = bindIntegrationType;

            Type menuManagerType = integrationAssembly.GetType("ExcelDna.Integration.MenuManager");
            addCommandMenu = menuManagerType.GetMethod("AddCommandMenu", BindingFlags.Static | BindingFlags.NonPublic);
            removeCommandMenus = menuManagerType.GetMethod("RemoveCommandMenus", BindingFlags.Static | BindingFlags.NonPublic);

            CreateUnhandledExceptionHandlerWrapper();
        }

        // We want to call an internal method in ExcelIntegration from a generator export wrapper.
        // In the past we could just hook up the MethodInfo, since we create a DynamicMethod.
        // We changed from DynamicMethod to a MethodBuilder to handle AccessViolations, but this puts us in a different
        // security context, where the restricted-visibility method calls are not allowed.
        // So we create a DynamicMethod here which makes the call fast, and expose internally via a public method.
        internal static void CreateUnhandledExceptionHandlerWrapper()
        {
			MethodInfo unhandledExceptionHandler = integrationType.GetMethod("HandleUnhandledException", BindingFlags.Static | BindingFlags.NonPublic);
            DynamicMethod ueh = new DynamicMethod("UnhandledExceptionHandler", typeof(object), new Type[] { typeof(object) }, true);
            ILGenerator uehIL = ueh.GetILGenerator();
            // uehIL.DeclareLocal(typeof(object));
            uehIL.Emit(OpCodes.Ldarg_0);
            uehIL.Emit(OpCodes.Call, unhandledExceptionHandler);
            uehIL.Emit(OpCodes.Ret);
            UnhandledExceptionHandler = (ExceptionHandler)ueh.CreateDelegate(typeof(ExceptionHandler));
        }

        internal static void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            addCommandMenu.Invoke(null, new object[] { commandName, menuName, menuText, description, shortCut, helpTopic});
        }

        internal static void RemoveCommandMenus()
        {
            removeCommandMenus.Invoke(null, null);
        }

        internal static void DnaLibraryAutoOpen()
        {
            integrationType.InvokeMember("DnaLibraryAutoOpen", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void DnaLibraryAutoClose()
        {
            integrationType.InvokeMember("DnaLibraryAutoClose", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        // TODO: Move this around a bit to clean up.
        // No longer called from xlAddInManagerInfo(12)
        internal static string DnaLibraryGetName()
        {
            return (string)integrationType.InvokeMember("DnaLibraryGetName", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static HRESULT DllRegisterServer()
        {
            return (HRESULT)integrationType.InvokeMember("DllRegisterServer", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static HRESULT DllUnregisterServer()
        {
            return (HRESULT)integrationType.InvokeMember("DllUnregisterServer", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        //internal static HRESULT DllGetClassObject(ref Guid clsid, ref Guid iid, out object ppunk)
        {
            HRESULT result;
            object[] args = new object[] {clsid, iid, null};
            result = (HRESULT)integrationType.InvokeMember("DllGetClassObject", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, args);
            ppunk = (IntPtr)args[2];
            return result;
        }

        internal static HRESULT DllCanUnloadNow()
        {
            return (HRESULT)integrationType.InvokeMember("DllCanUnloadNow", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void InitializeIntegration(string xllPath)
        {
            integrationType.InvokeMember("Initialize", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] {xllPath});        
        }

        internal static void DeInitializeIntegration()
        {
            integrationType.InvokeMember("DeInitialize", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void SyncMacro(double dValue)
        {
            integrationType.InvokeMember("SyncMacro", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, new object[] {dValue});
        }

        internal static void CalculationCanceled()
        {
            integrationType.InvokeMember("CalculationCanceled", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static void CalculationEnded()
        {
            integrationType.InvokeMember("CalculationEnded", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }

        internal static System.Diagnostics.TraceSource GetIntegrationTraceSource()
        {
            return (System.Diagnostics.TraceSource)integrationType.InvokeMember("GetIntegrationTraceSource", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.InvokeMethod, null, null, null);
        }
    }
}
