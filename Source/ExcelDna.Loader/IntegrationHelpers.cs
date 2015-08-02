//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Reflection;

namespace ExcelDna.Loader
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;

    // This class has a friend in XlCustomMarshal.
    // TODO: Nothing here is performance-critical, but the Invoke calls might be quite slow... use delegates?
    internal static class IntegrationHelpers
    {
        static Type integrationType;
        static MethodInfo addCommandMenu;
        static MethodInfo removeCommandMenus;
		internal static MethodInfo UnhandledExceptionHandler; // object->object
        
        public static void Bind(Assembly integrationAssembly, Type bindIntegrationType)
        {
            integrationType = bindIntegrationType;

            Type menuManagerType = integrationAssembly.GetType("ExcelDna.Integration.MenuManager");
            addCommandMenu = menuManagerType.GetMethod("AddCommandMenu", BindingFlags.Static | BindingFlags.NonPublic);
            removeCommandMenus = menuManagerType.GetMethod("RemoveCommandMenus", BindingFlags.Static | BindingFlags.NonPublic);

			UnhandledExceptionHandler = integrationType.GetMethod("HandleUnhandledException", BindingFlags.Static | BindingFlags.NonPublic);
        }

        public static void AddCommandMenu(string commandName, string menuName, string menuText, string description, string shortCut, string helpTopic)
        {
            addCommandMenu.Invoke(null, new object[] { commandName, menuName, menuText, description, shortCut, helpTopic});
        }

        public static void RemoveCommandMenus()
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
