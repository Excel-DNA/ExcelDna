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
    using System.Runtime.InteropServices.ComTypes;

    // The Excel-DNA .xll can also act as an in-process COM server.
    // This is implemented to support direct use of the RTD servers from the worksheet
    // using the =RTD(...) function.
    // TODO: Add explicit registration of types?
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
        public string TypeLibPath;

        public void RegisterServer()
        {
            // Register the ProgId for CLSIDFromProgID.
            string clsIdString = ClsId.ToString("B").ToUpperInvariant();
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\" + ProgId + @"\CLSID", null, clsIdString, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\InProcServer32", "ThreadingModel", "Both", RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\ProgID", null, ProgId, RegistryValueKind.String);

            if (!string.IsNullOrEmpty(TypeLibPath))
            {
                Guid? typeLibId = RegisterTypeLibrary();
                if (typeLibId.HasValue)
                {
                    Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\TypeLib", 
                        null, typeLibId.Value.ToString("B").ToUpperInvariant(), RegistryValueKind.String);
                }
            }
        }

        public void UnregisterServer()
        {
            if (!string.IsNullOrEmpty(TypeLibPath))
            {
                try
                {
                    UnregisterTypeLibrary();
                }
                catch (Exception e)
                {
                    Debug.Print("ComServer.UnregisterServer - UnregisterTypeLib error : " + e.ToString());
                }
            }
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
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\CLSID\" + ClsId.ToString("B").ToUpperInvariant());
            }
            catch (Exception e2)
            {
                Debug.Print("ComServer.UnregisterServer error : " + e2.ToString());
            }
        }

        public Guid? RegisterTypeLibrary()
        {
            ITypeLib typeLib;
            Guid libId;
            HRESULT hr = ComAPI.LoadTypeLib(TypeLibPath, out typeLib);
            if (hr != ComAPI.S_OK)
            {
                return null;
            }

            string helpDir = System.IO.Path.GetDirectoryName(TypeLibPath);
            if (!System.IO.Directory.Exists(helpDir))
            {
                helpDir = System.IO.Path.GetDirectoryName(DnaLibrary.XllPath);
            }

            // Deal with TYPELIBATTR
            IntPtr libAttrPtr;
            typeLib.GetLibAttr(out libAttrPtr);
            TYPELIBATTR typeLibAttr = (TYPELIBATTR)Marshal.PtrToStructure(libAttrPtr, typeof(TYPELIBATTR));

            libId = typeLibAttr.guid;
            string libIdString = libId.ToString("B").ToUpperInvariant();
            string version = typeLibAttr.wMajorVerNum.ToString() + "." + typeLibAttr.wMinorVerNum.ToString();
            
            // Get Friendly Name
            string friendlyName;
            string docString;
            int helpContext;
            string helpFile;
            typeLib.GetDocumentation(-1, out friendlyName, out docString, out helpContext, out helpFile);
            // string helpDir = System.IO.Path.GetDirectoryName(helpFile); // (or from TypeLibPath?)

            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\TypeLib\" + libIdString + @"\" + version, null, friendlyName, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\TypeLib\" + libIdString + @"\" + version + @"\" + "FLAGS", null, typeLibAttr.wLibFlags, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\TypeLib\" + libIdString + @"\" + version + @"\" + "HELPDIR", null, helpDir, RegistryValueKind.String);
            if (IntPtr.Size == 8)
            {
                Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\TypeLib\" + libIdString + @"\" + version + @"\" + typeLibAttr.lcid.ToString() + @"\win64", null, TypeLibPath, RegistryValueKind.String);
            }
            else
            {
                Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\TypeLib\" + libIdString + @"\" + version + @"\" + typeLibAttr.lcid.ToString() + @"\win32", null, TypeLibPath, RegistryValueKind.String);
            }

            typeLib.ReleaseTLibAttr(libAttrPtr);
            return libId;
        }

        public void UnregisterTypeLibrary()
        {
            try
            {
                ITypeLib typeLib;
                Guid libId;

                HRESULT hr = ComAPI.LoadTypeLib(TypeLibPath, out typeLib);
                if (hr != ComAPI.S_OK)
                {
                    return;
                }

                IntPtr libAttrPtr;
                typeLib.GetLibAttr(out libAttrPtr);
                TYPELIBATTR typeLibAttr = (TYPELIBATTR)Marshal.PtrToStructure(libAttrPtr, typeof(TYPELIBATTR));
                libId = typeLibAttr.guid;

                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\TypeLib\" + libId.ToString("B").ToUpperInvariant());

                typeLib.ReleaseTLibAttr(libAttrPtr);
                return;
            }
            catch (Exception e)
            {
                Debug.Print("TypeLibHelper.UnregisterServer error : " + e.ToString());
            }
        }

        //private bool _disposed;

        //public TypeLibHelper(string typeLibPath)
        //{
        //    _disposed = false;

        //}

        //public Guid LibId
        //{
        //    get;
        //    set;
        //}

        

        //private void Cleanup()
        //{
        //}

        //#region Disposable => Cleanup()
        //public void Dispose()
        //{
        //    Dispose(true);
        //    GC.SuppressFinalize(this);
        //}

        //protected virtual void Dispose(bool disposing)
        //{
        //    // Not thread-safe...
        //    if (!_disposed)
        //    {
        //        // if (disposing)
        //        // {
        //        //     // Here comes explicit free of other managed disposable objects.
        //        // }

        //        // Here comes clean-up
        //        Cleanup();
        //        _disposed = true;
        //    }
        //}

        //~TypeLibHelper()
        //{
        //    Dispose(false);
        //}
        //#endregion
    }
}
