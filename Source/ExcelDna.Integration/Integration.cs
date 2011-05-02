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
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration
{
    using HRESULT = System.Int32;
    using IID = System.Guid;
    using CLSID = System.Guid;
    using ExcelDna.ComInterop;

    // CAUTION: These functions are called _via reflection_ by
    // ExcelDna.Loader.XlLibrary to set up the link between the loader 
    // and the integration library.
    // Signatures, private/public etc. is fragile.

    internal delegate int TryExcelImplDelegate(int xlFunction, out object result, params object[] parameters);
    internal delegate void RegisterMethodsDelegate(List<MethodInfo> methods);
    internal delegate byte[] GetResourceBytesDelegate(string resourceName, int type); // types: 0 - Assembly, 1 - Dna file, 2 - Image
	public delegate object UnhandledExceptionHandler(object exceptionObject);

	// TODO: Rename to ExcelDnaAddIn? and make obsolete - type name should not be the same as namespace name.

    public static class Integration
    {
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

        // This is the only 'externally' exposed member.
        public static void RegisterMethods(List<MethodInfo> methods)
        {
            registerMethods(methods);
        }

		private static UnhandledExceptionHandler unhandledExceptionHandler;
		public static void RegisterUnhandledExceptionHandler(UnhandledExceptionHandler h)
		{
			unhandledExceptionHandler = h;
		}

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

        internal static void Initialize(string xllPath)
        {
			ExcelDnaUtil.Initialize();  // Set up window handle
            DnaLibrary.InitializeRootLibrary(xllPath);
        }

        internal static void DeInitialize()
        {
            DnaLibrary.DeInitialize();
        }

        internal static void DnaLibraryAutoOpen()
        {
			Debug.WriteLine("Enter Integration.DnaLibraryAutoOpen");
			try
			{
				DnaLibrary.CurrentLibrary.AutoOpen();
            }
			catch (Exception e)
			{
				Debug.WriteLine("Integration.DnaLibraryAutoOpen Exception: " + e);
			}
			Debug.WriteLine("Exit Integration.DnaLibraryAutoOpen");
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
            return ComServer.DllRegisterServer();
        }

        internal static HRESULT DllUnregisterServer()
        {
            return ComServer.DllUnregisterServer();
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
    }

    [Obsolete("Use ExcelDna.Integration.Integration class")]
    public class XlLibrary
    {
        [Obsolete("Use ExcelDna.Integration.Integration.RegisterMethods method")]
        public static void RegisterMethods(List<MethodInfo> methods)
        {
            Integration.RegisterMethods(methods);
        }
    }
}
