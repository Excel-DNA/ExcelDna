//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
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
    public static class ComServer
    {
        // COM Server support for persistently registered types.
        static readonly List<ExcelComClassType> registeredComClassTypes = new List<ExcelComClassType>();
        
        // Used for on-demand COM add-in (and RTD Server and CTP control) registration.
        // Added and removed from the Dictionary as we go along.
        static readonly Dictionary<CLSID, ComAPI.IClassFactory> registeredClassFactories = new Dictionary<CLSID, ComAPI.IClassFactory>();

        internal static void RegisterComClassTypes(List<ExcelComClassType> comClassTypes)
        {
            // Just merge registrations into the overall list.
            registeredComClassTypes.AddRange(comClassTypes);
        }

        internal static void RegisterComClassType(ExcelComClassType comClassType)
        {
            // Just merge registrations into the overall list.
            registeredComClassTypes.Add(comClassType);
        }

        internal static void RegisterClassFactory(CLSID clsId, ComAPI.IClassFactory classFactory)
        {
            // CONSIDER: Do we need to deal with the case where it is already in the list?
            //           For now we expect it to throw...
            registeredClassFactories.Add(clsId, classFactory);
        }

        internal static void UnregisterClassFactory(CLSID clsId)
        {
            registeredClassFactories.Remove(clsId);
        }

        // This may also be called by an add-in wanting to register programmatically (e.g. in AutoOpen)
        public static HRESULT DllRegisterServer()
        {
            foreach (ExcelComClassType comClass in registeredComClassTypes)
            {
                // TODO: Look for [ComRegisterFunction]
                comClass.RegisterServer();
            }
            return ComAPI.S_OK;
        }

        // This may also be called by an add-in wanting to unregister programmatically
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

            ComAPI.IClassFactory factory;
            if (registeredClassFactories.TryGetValue(clsid, out factory) ||
                TryGetComClassType(clsid, out factory))
            {
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

            // Otherwise it was not found
            ppunk = IntPtr.Zero;
            return ComAPI.CLASS_E_CLASSNOTAVAILABLE;
        }

        static bool TryGetComClassType(CLSID clsId, out ComAPI.IClassFactory factory)
        {
            // Check among the persistently registered classes
            foreach (ExcelComClassType comClass in registeredComClassTypes)
            {
                if (comClass.ClsId == clsId)
                {
                    factory = new ClassFactory(comClass);
                    return true;
                }
            }
            factory = null;
            return false;
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

        // Can throw UnauthorizedAccessException if nothing is writeable
        public void RegisterServer()
        {
            // Registering under the user key is problematic when Excel runs under an elevated account, e.g. when "Run as Administrator" 
            // or when UAC is disabled and the account is a member of local Adminstrators group.
            // In these cases the COM activation will ignore the user hive of the registry.
            // More info:
            // http://blogs.msdn.com/b/cjacks/archive/2007/02/21/per-user-com-registrations-and-elevated-processes-with-uac-on-windows-vista.aspx
            // and then changed in Vista SP1:
            // http://blogs.msdn.com/b/cjacks/archive/2008/06/06/per-user-com-registrations-and-elevated-processes-with-uac-on-windows-vista-sp1.aspx
            // http://blogs.msdn.com/b/cjacks/archive/2008/07/22/per-user-com-registrations-and-elevated-processes-with-uac-on-windows-vista-sp1-part-2-ole-automation.aspx

            string rootKeyName = RegistrationUtil.ClassesRootKey.Name;

            // Register the ProgId for CLSIDFromProgID.
            string clsIdString = ClsId.ToString("B").ToUpperInvariant();
            Registry.SetValue(rootKeyName + @"\" + ProgId + @"\CLSID", null, clsIdString, RegistryValueKind.String);
            Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
            Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\InProcServer32", "ThreadingModel", "Both", RegistryValueKind.String);
            Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\ProgID", null, ProgId, RegistryValueKind.String);

            if (!string.IsNullOrEmpty(TypeLibPath))
            {
                Guid? typeLibId = RegisterTypeLibrary(rootKeyName);
                if (typeLibId.HasValue)
                {
                    Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\TypeLib", 
                        null, typeLibId.Value.ToString("B").ToUpperInvariant(), RegistryValueKind.String);
                }
            }
        }
        
        // Can throw UnauthorizedAccessException if nothing is writeable
        public void UnregisterServer()
        {
            RegistryKey rootKey = RegistrationUtil.ClassesRootKey;

            if (!string.IsNullOrEmpty(TypeLibPath))
            {
                try
                {
                    UnregisterTypeLibrary(rootKey);
                }
                catch (Exception e)
                {
                    Debug.Print("ComServer.UnregisterServer - UnregisterTypeLib error : " + e.ToString());
                }
            }
            try
            {
                rootKey.DeleteSubKeyTree(ProgId);
            }
            catch (Exception e1)
            {
                Debug.Print("ComServer.UnregisterServer error : " + e1.ToString());
            }
            try
            {
                rootKey.DeleteSubKeyTree(@"CLSID\" + ClsId.ToString("B").ToUpperInvariant());
            }
            catch (Exception e2)
            {
                Debug.Print("ComServer.UnregisterServer error : " + e2.ToString());
            }
        }

        private Guid? RegisterTypeLibrary(string rootKeyName)
        {
            ITypeLib typeLib;
            Guid libId;
            HRESULT hr = ComAPI.LoadTypeLib(TypeLibPath, out typeLib);
            if (hr != ComAPI.S_OK)
            {
                return null;
            }

            string helpDir = System.IO.Path.GetDirectoryName(TypeLibPath);
            if (helpDir != null && !System.IO.Directory.Exists(helpDir))
            {
                helpDir = System.IO.Path.GetDirectoryName(DnaLibrary.XllPath);
            }

            // Deal with TYPELIBATTR
            IntPtr libAttrPtr;
            typeLib.GetLibAttr(out libAttrPtr);
            TYPELIBATTR typeLibAttr = (TYPELIBATTR)Marshal.PtrToStructure(libAttrPtr, typeof(TYPELIBATTR));

            libId = typeLibAttr.guid;
            string libIdString = libId.ToString("B").ToUpperInvariant();
            string version = typeLibAttr.wMajorVerNum.ToString(CultureInfo.InvariantCulture) + "." + typeLibAttr.wMinorVerNum.ToString(CultureInfo.InvariantCulture);
            
            // Get Friendly Name
            string friendlyName;
            string docString;
            int helpContext;
            string helpFile;
            typeLib.GetDocumentation(-1, out friendlyName, out docString, out helpContext, out helpFile);
            // string helpDir = System.IO.Path.GetDirectoryName(helpFile); // (or from TypeLibPath?)

            Registry.SetValue(rootKeyName + @"\TypeLib\" + libIdString + @"\" + version, null, friendlyName, RegistryValueKind.String);
            Registry.SetValue(rootKeyName + @"\TypeLib\" + libIdString + @"\" + version + @"\" + "FLAGS", null, typeLibAttr.wLibFlags, RegistryValueKind.DWord);
            if (helpDir != null)
            {
                Registry.SetValue(rootKeyName + @"\TypeLib\" + libIdString + @"\" + version + @"\" + "HELPDIR", null, helpDir, RegistryValueKind.String);
            }
            if (IntPtr.Size == 8)
            {
                Registry.SetValue(rootKeyName + @"\TypeLib\" + libIdString + @"\" + version + @"\" + typeLibAttr.lcid.ToString(CultureInfo.InvariantCulture) + @"\win64", null, TypeLibPath, RegistryValueKind.String);
            }
            else
            {
                Registry.SetValue(rootKeyName + @"\TypeLib\" + libIdString + @"\" + version + @"\" + typeLibAttr.lcid.ToString(CultureInfo.InvariantCulture) + @"\win32", null, TypeLibPath, RegistryValueKind.String);
            }

            typeLib.ReleaseTLibAttr(libAttrPtr);
            return libId;
        }

        private void UnregisterTypeLibrary(RegistryKey rootKey)
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

                rootKey.DeleteSubKeyTree(@"TypeLib\" + libId.ToString("B").ToUpperInvariant());

                typeLib.ReleaseTLibAttr(libAttrPtr);
            }
            catch (Exception e)
            {
                Debug.Print("TypeLibHelper.UnregisterServer error : " + e);
            }
        }

    }
}
