//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Configuration;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using Microsoft.Win32;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;

using CLSID = System.Guid;
using DWORD = System.Int32;
using HRESULT = System.Int32;
using IID = System.Guid;


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
        static RegistryKey _classesRootKey;
        static RegistryKey _clsIdRootKey;

        static RegistrationUtil()
        {
            Logger.ComAddIn.Verbose("Loading Ribbon/COM Add-In");
        }

        public static RegistryKey ClassesRootKey
        {
            get
            {
                if (_classesRootKey == null)
                {
                    // 3/22/2016: We use the intended hard coded reference of the HKCU hive to address the issue: https://groups.google.com/forum/#!topic/exceldna/CF_aNXTmV2Y
                    if (CanWriteMachineHive())
                    {
                        var subkey = @"Software\Classes";
                        Logger.ComAddIn.Verbose(@"RegistrationUtil.ClassesRootKey - Using HKLM\Software\Classes");
                        _classesRootKey = Registry.LocalMachine.CreateSubKey(subkey, RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }
                    else if (CanWriteUserHive())
                    {
                        string subkey = WindowsIdentity.GetCurrent().User.ToString() + @"_CLASSES";
                        Logger.ComAddIn.Verbose("RegistrationUtil.ClassesRootKey - Using Users subkey {0}", subkey);
                        _classesRootKey = Registry.Users.CreateSubKey(subkey, RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }
                    else
                    {
                        // We have no further plan
                        Logger.ComAddIn.Error("RegistrationUtil - Unable to write to Machine or Users hives of registry - Ribbon/COM Add-In load cancelled");
                        throw new UnauthorizedAccessException("RegistrationUtil - Unable to write to Machine or Users hives of registry");
                    }
                }
                return _classesRootKey;
            }
        }

        public static RegistryKey ClsIdRootKey
        {
            get
            {
                if (_clsIdRootKey == null)
                {
                    _clsIdRootKey = ClassesRootKey.CreateSubKey("CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
                return _clsIdRootKey;
            }
        }

        static bool CanWriteMachineHive()
        {
            // This is not an easy question to answer, due to Registry Virtualization: http://msdn.microsoft.com/en-us/library/aa965884(v=vs.85).aspx
            // So if registry virtualization is active, the machine writes will redirect to a special user key.
            // I don't know how to detect that case, so we'll just write to the virtualized location.
            string machineClassesRoot = @"Software\Classes";
            const string testKeyName = "_ExcelDna.PermissionsTest";
            try
            {
                RegistryKey classesKey = Registry.LocalMachine.CreateSubKey(machineClassesRoot, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (classesKey == null)
                {
                    Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteMachineHive - Opening LocalMachineClassesRoot as ReadWrite failed - returning False");
                    return false;
                }

                RegistryKey testKey = classesKey.CreateSubKey(testKeyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (testKey == null)
                {
                    Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteMachineHive - Creating test sub key failed - returning False");
                    return false;
                }

                classesKey.DeleteSubKeyTree(testKeyName);
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteMachineHive - returning True");

                // Looks fine, even though it might well be virtualized to some part of the user hive.
                // I'd have preferred to return false in the virtualized case, but don't know how to detect it.
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteMachineHive - UnauthorizedAccessException - False");
                return false;
            }
            catch (SecurityException)
            {
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteMachineHive - SecurityException - False");
                return false;
            }
            catch (Exception e)
            {
                Logger.ComAddIn.Error(e, "RegistrationUtil.CanWriteMachineHive - Unexpected exception - False");
                return false;
            }
        }

        static bool CanWriteUserHive()
        {
            string userClassesRoot = WindowsIdentity.GetCurrent().User.ToString() + @"_CLASSES";
            const string testKeyName = "_ExcelDna.PermissionsTest";
            try
            {
                RegistryKey classesKey = Registry.Users.CreateSubKey(userClassesRoot, RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (classesKey == null)
                {
                    Logger.ComAddIn.Error("RegistrationUtil.CanWriteUserHive - Opening UserClassesRoot - Unexpected failure - False");
                    return false;
                }

                RegistryKey testKey = classesKey.CreateSubKey(testKeyName);
                if (testKey == null)
                {
                    Logger.ComAddIn.Error("RegistrationUtil.CanWriteUserHive - Creating test sub key - Unexpected failure - False");
                    return false;
                }

                classesKey.DeleteSubKeyTree(testKeyName);
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteUserHive - True");

                // Looks fine, even though it might well be virtualized to some part of the user hive.
                // I'd have preferred to return false in the virtualized case, but don't know how to detect it.
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteUserHive - UnauthorizedAccessException - False");
                return false;
            }
            catch (SecurityException)
            {
                Logger.ComAddIn.Verbose("RegistrationUtil.CanWriteUserHive - SecurityException - False");
                return false;
            }
            catch (Exception e)
            {
                Logger.ComAddIn.Error(e, "RegistrationUtil.CanWriteUserHive - Unexpected exception - False");
                return false;
            }
        }

        // Registry calls 
        public static RegistryKey UsersCreateSubKey(string subkey, RegistryKeyPermissionCheck permissionsCheck)
        {
            Logger.ComAddIn.Verbose("RegistrationUtil.UsersCreateSubKey({0}, {1})", subkey, permissionsCheck);
            return Registry.Users.CreateSubKey(subkey, permissionsCheck);
        }

        public static void UsersDeleteSubKey(string subkey)
        {
            Logger.ComAddIn.Verbose("RegistrationUtil.UsersDeleteSubKey({0})", subkey);
            Registry.Users.DeleteSubKey(subkey);
        }

        public static void KeySetValue(RegistryKey key, string name, object value, RegistryValueKind valueKind)
        {
            Logger.ComAddIn.Verbose("RegistrationUtil.KeySetValue({0}, {1}, {2}, {3})", key.Name, name, value.ToString(), valueKind.ToString());
            key.SetValue(name, value, valueKind);
        }

        public static void DeleteSubKeyTree(RegistryKey key, string subkey)
        {
            Logger.ComAddIn.Verbose("RegistrationUtil.DeleteSubKeyTree({0}, {1})", key.Name, subkey);
            key.DeleteSubKeyTree(subkey);
        }

        public static void SetValue(string keyName, string valueName, object value, RegistryValueKind valueKind)
        {
            Logger.ComAddIn.Verbose("RegistrationUtil.SetValue({0}, {1}, {2}, {3})", keyName, valueName, value.ToString(), valueKind.ToString());
            Registry.SetValue(keyName, valueName, value, valueKind);
        }

        // Helper for AppSettings (can move somewhere later)
        static bool AppSettingsFlag(string key)
        {
            var value = ConfigurationManager.AppSettings[key];
            if (value == null)
                return false;
            
            bool flag;
            if (bool.TryParse(value, out flag))
                return flag;

            return false;
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
        readonly string _subKeyName;

        public ComAddInRegistration(string progId, string friendlyName, string description)
        {
            // Register the ProgId as a COM Add-In in Excel.
            _progId = progId;
            _subKeyName = WindowsIdentity.GetCurrent().User.ToString() + @"\Software\Microsoft\Office\Excel\Addins\" + progId;
            Logger.ComAddIn.Verbose("ComAddInRegistration - Creating User SubKey: " + _subKeyName);

            RegistryKey rk = RegistrationUtil.UsersCreateSubKey(_subKeyName, RegistryKeyPermissionCheck.ReadWriteSubTree);
            RegistrationUtil.KeySetValue(rk, "LoadBehavior", 0, RegistryValueKind.DWord);
            RegistrationUtil.KeySetValue(rk, "FriendlyName", friendlyName, RegistryValueKind.String);
            RegistrationUtil.KeySetValue(rk, "Description", description, RegistryValueKind.String);
        }

        protected override void Deregister()
        {
            // Remove Add-In registration from Excel
            Logger.ComAddIn.Verbose("ComAddInRegistration - Deleting User SubKey: " + _subKeyName);
            RegistrationUtil.UsersDeleteSubKey(_subKeyName);
        }
    }

    internal class ProgIdRegistration : Registration
    {
        readonly string _progId;

        public ProgIdRegistration(string progId, CLSID clsId)
        {
            _progId = progId;
            string rootKeyName = RegistrationUtil.ClassesRootKey.Name;
            string progIdKeyName = rootKeyName + @"\" + _progId;
            string value = clsId.ToString("B").ToUpperInvariant();
            // Register the ProgId for CLSIDFromProgID.
            Logger.ComAddIn.Verbose("ProgIdRegistration - Set Value - {0} -> {1}", progIdKeyName, value);
            RegistrationUtil.SetValue(progIdKeyName + @"\CLSID", null, value, RegistryValueKind.String);
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            Logger.ComAddIn.Verbose("ProgIdRegistration - Delete SubKey {0}", _progId);
            RegistrationUtil.DeleteSubKeyTree(RegistrationUtil.ClassesRootKey, _progId);
        }
    }

    internal class ClsIdRegistration : Registration
    {
        readonly Guid _clsId;
        readonly string _clsIdString;

        public ClsIdRegistration(CLSID clsId, string progId)
        {
            _clsId = clsId;
            _clsIdString = clsId.ToString("B").ToUpperInvariant();
            string clsIdRootKeyName = RegistrationUtil.ClsIdRootKey.Name;
            string clsIdKeyName = clsIdRootKeyName + "\\" + _clsIdString;
            // Register the CLSID

            // NOTE: Remember that all the CLSID keys are redirected under WOW64.
            Logger.ComAddIn.Verbose("ClsIdRegistration - Set Values - {0} ({1}) - {2}", clsIdKeyName, progId, DnaLibrary.XllPath);
            RegistrationUtil.SetValue(clsIdKeyName + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
            RegistrationUtil.SetValue(clsIdKeyName + @"\InProcServer32", "ThreadingModel", "Both", RegistryValueKind.String);
            if (!string.IsNullOrEmpty(progId))
            {
                RegistrationUtil.SetValue(clsIdKeyName + @"\ProgID", null, progId, RegistryValueKind.String);
            }
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            Logger.ComAddIn.Verbose("ClsIdRegistration - Delete SubKey {0}", RegistrationUtil.ClsIdRootKey.Name + _clsIdString);
            RegistrationUtil.DeleteSubKeyTree(RegistrationUtil.ClsIdRootKey, _clsIdString);
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

    // We want to temporarily set Application.AutomationSecurity = 1 (msoAutomationSecurityLow)
    internal class AutomationSecurityOverride : Registration
    {
        object _app;
        object _oldValue = null;
        public AutomationSecurityOverride(object app)
        {
            _app = app;
            _oldValue = SetAutomationSecurity(1);
        }

        protected override void Deregister()
        {
            if (_oldValue != null)
            {
                SetAutomationSecurity(_oldValue);
            }
        }

        object SetAutomationSecurity(object value)
        {
            CultureInfo ci = new CultureInfo(1033);
            Type appType = _app.GetType();
            try
            {
                var oldValue = appType.InvokeMember("AutomationSecurity", BindingFlags.GetProperty, null, _app, null, ci);
                if (oldValue.Equals(value)) // Careful...they're boxed ints
                    return null;

                Logger.ComAddIn.Verbose("AutomationSecurityEnable - Setting Application.AutomationSecurity to {0}", value);
                appType.InvokeMember("AutomationSecurity", BindingFlags.SetProperty, null, _app, new object[] { value }, ci);
                return oldValue;
            }
            catch (Exception ex)
            {
                // We're not going to treat this as an error - we expect the COM add-in load to fail, which may pop up an error.
                Logger.ComAddIn.Info("AutomationSecurityEnable - SetAutomationSecurity error: {0}", ex.ToString());
            }
            return null;
        }
    }

    #endregion

}
