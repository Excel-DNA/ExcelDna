using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using ExcelDna.Integration;
using ExcelDna.ComInterop.ComRegistration;

namespace ExcelDna.ComInterop
{
    using HRESULT = System.Int32;
    using IID = System.Guid;
    using CLSID = System.Guid;

    // The Excel-DNA .xll can also act as an in-process COM server.
    // This is implemented to support direct use of the RTD servers from the worksheet
    // using the =RTD(...) function.
    // TODO: Add explicit registration of types.
    // TODO: Add on-demand registration.
    public class ComServer
    {
        // Internal COM Server support.
        static List<ExcelComClassType> registeredComClassTypes = new List<ExcelComClassType>();

        internal static void RegisterComClassTypes(List<ExcelComClassType> comClassTypes)
        {
            // Just merge registrations into the overall list.
            registeredComClassTypes.AddRange(comClassTypes);
        }

        // This may also be called by an add-in wanting to register
        // CONSIDER: Should this rather use RegistrationServices class?
        public static HRESULT DllRegisterServer()
        {
            foreach (ExcelComClassType comClass in registeredComClassTypes)
            {
                // TODO: Look for [ComRegisterFunction]
                comClass.RegisterServer();
            }
            return ComAPI.S_OK;
        }

        // This may also be called by an add-in wanting to unregister
        public static HRESULT DllUnregisterServer()
        {
            foreach (ExcelComClassType comClass in registeredComClassTypes)
            {
                comClass.UnregisterServer();
            }
            return ComAPI.S_OK;
        }

        internal static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        {
            if (iid != ComAPI.guidIClassFactory)
            {
                ppunk = IntPtr.Zero;
                return ComAPI.E_INVALIDARG;
            }
            foreach (ExcelComClassType comClass in registeredComClassTypes)
            {
               if (comClass.ClsId == clsid)
               {
                   ClassFactory factory = new ClassFactory(comClass);
                   IntPtr punkFactory = Marshal.GetIUnknownForObject(factory);
                   HRESULT hrQI = Marshal.QueryInterface(punkFactory, ref iid, out ppunk);
                   Marshal.Release(punkFactory);
                   if (hrQI == ComAPI.S_OK)
                   {
                       return ComAPI.S_OK;
                   }
                   else
                   {
                       return ComAPI.E_UNEXPECTED;
                   }
               }
            }
            ppunk = IntPtr.Zero;
            return ComAPI.CLASS_E_CLASSNOTAVAILABLE;
        }

        internal static HRESULT DllCanUnloadNow()
        {
            // CONSIDER: Allow unloading - but how to keep track of this.....?
            return ComAPI.S_FALSE;
        }

    }

    internal class ExcelComClassType
    {
        public Guid ClsId;
        public string ProgId;
        public Type Type;
        public bool IsRtdServer;

        public void RegisterServer()
        {
            // Register the ProgId for CLSIDFromProgID.
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\" + ProgId + @"\CLSID", null, ClsId.ToString("B"), RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + ClsId.ToString("B") + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
        }

        public void UnregisterServer()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\" + ProgId);
            }
            catch (Exception e1)
            {
                Debug.Print("ComServer.UnregisterServer error : " + e1.ToString());
            }
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\CLSID\" + ClsId.ToString("B"));
            }
            catch (Exception e2)
            {
                Debug.Print("ComServer.UnregisterServer error : " + e2.ToString());
            }
        }
    }
}
