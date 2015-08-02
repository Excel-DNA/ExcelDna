//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Win32;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.Rtd;

using CLSID     = System.Guid;
using DWORD     = System.Int32;
using HRESULT   = System.Int32;
using IID       = System.Guid;
using ExcelDna.Logging;

namespace ExcelDna.ComInterop.ComRegistration
{
    // This implements a COM class factory for the given type
    // with some customization to allow wrapping of Rtd servers.
    // Does not work with the just-in-time registration into the user's hive, when running under elevated UAC token.
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    internal class ClassFactory : ComAPI.IClassFactory
    {
        private ExcelComClassType _comClass;

        public ClassFactory(Type type)
        {
            _comClass = new ExcelComClassType 
            {
                Type = type,
                IsRtdServer = false
            };
        }

        public ClassFactory(ExcelComClassType excelComClassType)
        {
            _comClass = excelComClassType;
        }

        public HRESULT CreateInstance([In] IntPtr pUnkOuter, [In] ref IID riid, [Out] out IntPtr ppvObject)
        {
            // Suspend the C API (helps to prevent some Excel-crashing scenarios)
            using (XlCall.Suspend())
            {
                ppvObject = IntPtr.Zero;
                object instance = Activator.CreateInstance(_comClass.Type);

                // If not an ExcelRtdServer, create safe wrapper that also maps types.
                if (_comClass.IsRtdServer && !instance.GetType().IsSubclassOf(typeof(ExcelRtdServer)))
                {
                    // wrap instance in RtdWrapper
                    RtdServerWrapper rtdServerWrapper = new RtdServerWrapper(instance, _comClass.ProgId);
                    instance = rtdServerWrapper;
                }

                if (pUnkOuter != IntPtr.Zero)
                {
                    // For now no aggregation support - could do Marshal.CreateAggregatedObject?
                    return ComAPI.CLASS_E_NOAGGREGATION;
                }
                if (riid == ComAPI.guidIUnknown)
                {
                    ppvObject = Marshal.GetIUnknownForObject(instance);
                }
                else
                {
                    ppvObject = Marshal.GetIUnknownForObject(instance);
                    HRESULT hrQI = Marshal.QueryInterface(ppvObject, ref riid, out ppvObject);
                    Marshal.Release(ppvObject);
                    if (hrQI != ComAPI.S_OK)
                    {
                        return ComAPI.E_NOINTERFACE;
                    }
                }
                return ComAPI.S_OK;
            }
        }

        public int LockServer(bool fLock)
        {
            return ComAPI.S_OK;
        }
    }

    // This is a class factory that serves as a singleton 'factory' for a given object
    // - it will return exactly that object when CreateInstance is called 
    // (checking interface support).
    // Used for the RTD classes.
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    internal class SingletonClassFactory : ComAPI.IClassFactory
    {
        private object _instance;

        public SingletonClassFactory(object instance)
        {
            _instance = instance;
        }

        public HRESULT CreateInstance([In] IntPtr pUnkOuter, [In] ref IID riid, [Out] out IntPtr ppvObject)
        {
            using (XlCall.Suspend())
            {
                ppvObject = IntPtr.Zero;
                if (pUnkOuter != IntPtr.Zero)
                {
                    // For now no aggregation support - could do Marshal.CreateAggregatedObject?
                    return ComAPI.CLASS_E_NOAGGREGATION;
                }
                if (riid == ComAPI.guidIUnknown)
                {
                    ppvObject = Marshal.GetIUnknownForObject(_instance);
                }
                else if (riid == ComAPI.guidIDTExtensibility2)
                {
                    ppvObject = Marshal.GetComInterfaceForObject(_instance, typeof(IDTExtensibility2));
                }
                else if (riid == ComAPI.guidIRtdServer)
                {
                    ppvObject = Marshal.GetComInterfaceForObject(_instance, typeof(IRtdServer));
                }
                else // Unsupported interface for us.
                {
                    return ComAPI.E_NOINTERFACE;
                }
                return ComAPI.S_OK;
            }
        }

        public int LockServer(bool fLock)
        {
            return ComAPI.S_OK;
        }
    }

    #region Registration Helpers

    // We check whether the machine-hive HKCR is writeable (by attempting a write)
    // If it is writeable, we register to the machine hive. Otherwise we fall back to the user hive.
    // Some care is needed here. See notes at ComServer.RegisterServer().

    internal static class RegistrationUtil
    {
        static RegistryKey _rootKey;

        public static RegistryKey ClassesRootKey
        {
            get
            {
                if (_rootKey == null)
                {
                    _rootKey = CanWriteMachineHive() ? 
                                Registry.ClassesRoot : 
                                Registry.CurrentUser.CreateSubKey(@"Software\Classes", RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
                return _rootKey;
            }
        }

        static bool CanWriteMachineHive()
        {
            // This is not an easy question to answer, due to Registry Virtualization: http://msdn.microsoft.com/en-us/library/aa965884(v=vs.85).aspx
            // So if registry virtualization is active, the machine writes will redirect to a special user key.
            // I don't know how to detect that case, so we'll just write to the virtualized location.

            const string testKeyName = "_ExcelDna.PermissionsTest";
            try
            {
                RegistryKey testKey = Registry.ClassesRoot.CreateSubKey(testKeyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (testKey == null)
                {
                    Logger.ComAddIn.Error("Unexpected failure in CanWriteMachineHive check");
                    return false;
                }
                else
                {
                    Registry.ClassesRoot.DeleteSubKeyTree(testKeyName);

                    // Looks fine, even though it might well be virtualized to some part of the user hive.
                    // I'd have preferred to return false in the virtualized case, but don't know how to detect it.
                    return true;
                }
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (SecurityException)
            {
                return false;
            }
            catch (Exception e)
            {
                Logger.ComAddIn.Error(e, "Unexpected exception in CanWriteMachineHive check");
                return false;
            }
        }
    }

    // Disposable base class
    internal abstract class Registration : IDisposable
    {
        private bool _disposed;

        public Registration()
        {
            _disposed = false;
        }

        protected abstract void Deregister();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Not thread-safe...
            if (!_disposed)
            {
                // if (disposing)
                // {
                //     // Here comes explicit free of other managed disposable objects.
                // }

                // Here comes clean-up
                try
                {
                    Deregister();
                }
                catch
                {
                    // Ignore exception here - we've tried our best to clean up.
                    // CONSIDER: Might be useful to log this error?
                }
                _disposed = true;
            }
        }

        ~Registration()
        {
            Dispose(false);
        }

    }

    internal class ComAddInRegistration : Registration
    {
        
        readonly string _progId;

        public ComAddInRegistration(string progId, string friendlyName, string description)
        {
            _progId = progId;
            // Register the ProgId as a COM Add-In in Excel.
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\" + progId, "LoadBehavior", 0, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\" + progId, "FriendlyName", friendlyName, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\" + progId, "Description", description, RegistryValueKind.String);
        }

        protected override void Deregister()
        {
            // Remove Add-In registration from Excel
            Registry.CurrentUser.DeleteSubKey(@"Software\Microsoft\Office\Excel\Addins\" + _progId);
        }
    }

    internal class ProgIdRegistration : Registration
    {
        readonly string _progId;

        public ProgIdRegistration(string progId, CLSID clsId)
        {
            _progId = progId;
            string rootKeyName = RegistrationUtil.ClassesRootKey.Name;

            // Register the ProgId for CLSIDFromProgID.
            Registry.SetValue(rootKeyName + @"\" + _progId + @"\CLSID", null, clsId.ToString("B").ToUpperInvariant(), RegistryValueKind.String);
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            RegistrationUtil.ClassesRootKey.DeleteSubKeyTree(_progId);
        }
    }

    internal class ClsIdRegistration : Registration
    {
        readonly Guid _clsId;
        public ClsIdRegistration(CLSID clsId, string progId)
        {
            _clsId = clsId;
            string clsIdString = clsId.ToString("B").ToUpperInvariant();
            string rootKeyName = RegistrationUtil.ClassesRootKey.Name;

            // Register the CLSID

            // NOTE: Remember that all the CLSID keys are redirected under WOW64.

            Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
            Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\InProcServer32", "ThreadingModel", "Both", RegistryValueKind.String);
            if (!string.IsNullOrEmpty(progId))
            {
                Registry.SetValue(rootKeyName + @"\CLSID\" + clsIdString + @"\ProgID", null, progId, RegistryValueKind.String);
            }
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            RegistrationUtil.ClassesRootKey.DeleteSubKeyTree(@"CLSID\" + _clsId.ToString("B").ToUpperInvariant());
        }
    }

    // Implements the IClassFactory factory registration by (temporarily) adding it to the list of classes implemented through the ComServer.
    // We then write the registry keys that will point Excel to the .xll as the provider of this class, 
    // and the ComServer will handle the DllGetClassObject call, returning the IClassFactory.
    internal class SingletonClassFactoryRegistration : Registration
    {
        CLSID _clsId;
        public SingletonClassFactoryRegistration(object instance, CLSID clsId)
        {
            _clsId = clsId;
            SingletonClassFactory factory = new SingletonClassFactory(instance);
            ComServer.RegisterClassFactory(_clsId, factory);
        }

        protected override void Deregister()
        {
            ComServer.UnregisterClassFactory(_clsId);
        }
    }

    #endregion
}