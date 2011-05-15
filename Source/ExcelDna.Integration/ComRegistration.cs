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
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Drawing;
using Microsoft.Win32;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;

using HRESULT = System.Int32;
using IID = System.Guid;
using CLSID = System.Guid;
using DWORD = System.UInt32;

namespace ExcelDna.ComInterop.ComRegistration
{
    // This implements a COM class factory for the given type
    // with some customization to allow wrapping of Rtd servers.
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
            ppvObject = IntPtr.Zero;
            object instance = Activator.CreateInstance(_comClass.Type);

            if (_comClass.IsRtdServer)
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

        public int LockServer(bool fLock)
        {
            //Debug.Fail("LockServer not implemented yet....?");
            //throw new NotImplementedException();
            return ComAPI.S_OK;
        }
    }

    // This is a class factory that serve as a singleton 'factory' for a given object
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

        public int LockServer(bool fLock)
        {
            //Debug.Fail("LockServer not implemented yet....?");
            //throw new NotImplementedException();
            return ComAPI.S_OK;
        }
    }

    #region Registration Helpers

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
                Deregister();
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
        private string _progId;

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
        public string ProgId
        {
            get;
            private set;
        }

        public ProgIdRegistration(string progId, CLSID clsId)
        {
            ProgId = progId;
            // Register the ProgId for CLSIDFromProgID.
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\" + ProgId + @"\CLSID", null, clsId.ToString("B").ToUpperInvariant(), RegistryValueKind.String);
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\" + ProgId);
        }
    }

    internal class ClsIdRegistration : Registration
    {
        public Guid ClsId
        {
            get;
            private set;
        }

        public ClsIdRegistration(CLSID clsId, string progId)
        {
            ClsId = clsId;
            string clsIdString = clsId.ToString("B").ToUpperInvariant();
            // Register the CLSID
            //Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsId.ToString("B"), null, "Excel RTD Helper Class", RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\InProcServer32", null, DnaLibrary.XllPath, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\InProcServer32", "ThreadingModel", "Both", RegistryValueKind.String);
            if (!string.IsNullOrEmpty(progId))
            {
                Registry.SetValue(@"HKEY_CURRENT_USER\Software\Classes\CLSID\" + clsIdString + @"\ProgID", null, progId, RegistryValueKind.String);
            }
        }

        protected override void Deregister()
        {
            // Deregister the ProgId for CLSIDFromProgID.
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\CLSID\" + ClsId.ToString("B").ToUpperInvariant());
        }
    }

    internal class ClassFactoryRegistration : Registration
    {
        private const DWORD CLSCTX_INPROC_SERVER = 0x1;
        private const DWORD REGCLS_SINGLEUSE = 0;
        private const DWORD REGCLS_MULTIPLEUSE = 1;

        private DWORD _classRegister;

        public ClassFactoryRegistration(Type type, CLSID clsId)
        {
            ClassFactory factory = new ClassFactory(type);
            IntPtr pFactory = Marshal.GetIUnknownForObject(factory);
            HRESULT result = ComAPI.CoRegisterClassObject(ref clsId, pFactory,
                                CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE, out _classRegister);

            if (result != ComAPI.S_OK)
            {
                throw new InvalidOperationException("CoRegisterClassObject failed.");
            }
        }

        protected override void Deregister()
        {
            if (_classRegister != 0)
            {
                HRESULT result = ComAPI.CoRevokeClassObject(_classRegister);
                if (result != ComAPI.S_OK)
                {
                    Debug.Print("ClassFactory deregistration failed. Result: {0}", result);
                }
                _classRegister = 0;
            }
        }
    }

    internal class SingletonClassFactoryRegistration : Registration
    {
        private const DWORD CLSCTX_INPROC_SERVER = 0x1;
        private const DWORD REGCLS_SINGLEUSE = 0;
        private const DWORD REGCLS_MULTIPLEUSE = 1;

        private object _instance;
        private DWORD _classRegister;

        public SingletonClassFactoryRegistration(object instance, CLSID clsId)
        {
            _instance = instance;
            SingletonClassFactory factory = new SingletonClassFactory(instance);
            IntPtr pFactory = Marshal.GetIUnknownForObject(factory);
            // In versions < 0.29 we registered as REGCLS_SINGLEUSE even though it is not supposed to work for inproc servers.
            // It seems to do no harm to keep the ClassObject around.
            HRESULT result = ComAPI.CoRegisterClassObject(ref clsId, pFactory,
                                CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE, out _classRegister);

            if (result != ComAPI.S_OK)
            {
                throw new InvalidOperationException("CoRegisterClassObject failed.");
            }
        }

        protected override void Deregister()
        {
            if (_classRegister != 0)
            {
                HRESULT result = ComAPI.CoRevokeClassObject(_classRegister);
                if (result != ComAPI.S_OK)
                {
                    Debug.Print("ClassFactory deregistration failed. Result: {0}", result);
                }
                _classRegister = 0;
            }
        }
    }
    #endregion
}