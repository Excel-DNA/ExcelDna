
//// from www.danielmoth.com/Blog
//// Not for use on machines that you value.
//// Just a proof of concept.

//using System;
//using System.Reflection;
//using System.Runtime.InteropServices;
//using Microsoft.Win32;
//using Microsoft.Win32.SafeHandles;

//public static class RegistryHelper
//{
//  public static RegistryKey MothCreateVolatileSubKey(RegistryKey rk, string subkey, RegistryKeyPermissionCheck permissionCheck)
//  {
//    Type RK2 = rk.GetType();
//    BindingFlags bfStatic = BindingFlags.NonPublic | BindingFlags.Static;
//    BindingFlags bfInstance = BindingFlags.NonPublic | BindingFlags.Instance;
//    RK2.GetMethod("ValidateKeyName", bfStatic).Invoke(null, new object[] { subkey });
//    RK2.GetMethod("ValidateKeyMode", bfStatic).Invoke(null, new object[] { permissionCheck });
//    RK2.GetMethod("EnsureWriteable", bfInstance).Invoke(rk, null);
//    subkey = (string)RK2.GetMethod("FixupName", bfStatic).Invoke(null, new object[] { subkey });
//    if (!(bool)RK2.GetField("remoteKey", bfInstance).GetValue(rk))
//    {
//      RegistryKey key = (RegistryKey)RK2.GetMethod("InternalOpenSubKey", bfInstance, null, new Type[] { typeof(string), typeof(bool) }, null).Invoke(rk, new object[] { subkey, permissionCheck != RegistryKeyPermissionCheck.ReadSubTree });
//      if (key != null)
//      {
//        RK2.GetMethod("CheckSubKeyWritePermission", bfInstance).Invoke(rk, new object[] { subkey });
//        RK2.GetMethod("CheckSubTreePermission", bfInstance).Invoke(rk, new object[] { subkey, permissionCheck });
//        RK2.GetField("checkMode", bfInstance).SetValue(key, permissionCheck);
//        return key;
//      }
//    }
//    RK2.GetMethod("CheckSubKeyCreatePermission", bfInstance).Invoke(rk, new object[] { subkey });
//    int lpdwDisposition = 0;
//    IntPtr hkResult;
//    Type srh = Type.GetType("Microsoft.Win32.SafeHandles.SafeRegistryHandle");
//    object temp = RK2.GetField("hkey", bfInstance).GetValue(rk);
//    int getregistrykeyaccess;
//    SafeHandleZeroOrMinusOneIsInvalid rkhkey = (SafeHandleZeroOrMinusOneIsInvalid)temp;
//    getregistrykeyaccess = (int)RK2.GetMethod("GetRegistryKeyAccess", bfStatic, null, new Type[] { typeof(bool) }, null).Invoke(null, new object[] { permissionCheck != RegistryKeyPermissionCheck.ReadSubTree });
//    int errorCode = RegCreateKeyEx(rkhkey, subkey, 0, null, 1, getregistrykeyaccess, IntPtr.Zero, out hkResult, out lpdwDisposition);
//    string rkkeyName = (string)RK2.GetField("keyName", bfInstance).GetValue(rk);
//    if (errorCode == 0 && hkResult.ToInt32() > 0)
//    {
//      bool rkremoteKey = (bool)RK2.GetField("remoteKey", bfInstance).GetValue(rk);
//      object hkResult2 = srh.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new Type[] { typeof(IntPtr), typeof(bool) }, null).Invoke(new object[] { hkResult, true });
//      RegistryKey key2 = (RegistryKey)RK2.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new Type[] { hkResult2.GetType(), typeof(bool), typeof(bool), typeof(bool), typeof(bool) }, null).Invoke(new object[] { hkResult2, permissionCheck != RegistryKeyPermissionCheck.ReadSubTree, false, rkremoteKey, false });
//      RK2.GetMethod("CheckSubTreePermission", bfInstance).Invoke(rk, new object[] { subkey, permissionCheck });
//      RK2.GetField("checkMode", bfInstance).SetValue(key2, permissionCheck);
//      if (subkey.Length == 0)
//        RK2.GetField("keyName", bfInstance).SetValue(key2, rkkeyName);
//      else
//        RK2.GetField("keyName", bfInstance).SetValue(key2, rkkeyName + @"\" + subkey);
//      key2.Close();
//      return rk.OpenSubKey(subkey, true);
//    }
//    if (errorCode != 0)
//      RK2.GetMethod("Win32Error", bfInstance).Invoke(rk, new object[] { errorCode, rkkeyName + @"\" + subkey });
//    return null;
//  }

//  [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
//  private static extern int RegCreateKeyEx(SafeHandleZeroOrMinusOneIsInvalid hKey, string lpSubKey, int Reserved, string lpClass, int dwOptions, int samDesigner, IntPtr lpSecurityAttributes, out IntPtr hkResult, out int lpdwDisposition);
//}


// This file from fandrei
// https://github.com/fandrei/ExcelDnaExperimenting/blob/48d84398bb381d6d0e13a03346449293dbc386f0/Source/ExcelDna.Integration/Utils/RegistryHelper.cs
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace ExcelDna.Integration.Utils
{
    internal static class RegistryHelper
    {
        public static RegistryKey CreateVolatileSubKey(RegistryKey baseKey, string subKeyName, RegistryKeyPermissionCheck permissionCheck)
        {
            var keyType = baseKey.GetType();

            var handleField = keyType.GetField("hkey", BindingFlags.NonPublic | BindingFlags.Instance);
            var baseKeyHandle = (SafeHandleZeroOrMinusOneIsInvalid)(handleField.GetValue(baseKey));

            var accessRights = (int)(RegSAM.QueryValue);
            if (permissionCheck == RegistryKeyPermissionCheck.ReadWriteSubTree)
                accessRights = (int)(RegSAM.QueryValue | RegSAM.CreateSubKey | RegSAM.SetValue);

            int lpdwDisposition;
            var hkResult = IntPtr.Zero;

            try
            {
                var errorCode = RegCreateKeyEx(baseKeyHandle, subKeyName, 0, null, (int)RegOption.Volatile,
                    accessRights, IntPtr.Zero, out hkResult, out lpdwDisposition);

                if (errorCode == 5)
                    throw new UnauthorizedAccessException();

                if (errorCode != 0 || hkResult.ToInt32() <= 0)
                    throw new Win32Exception();
            }
            finally
            {
                // NOTE: This does not look right...
                //if (hkResult.ToInt32() > 0)
                //    Marshal.Release(hkResult);
            }

            return baseKey.OpenSubKey(subKeyName, true);
        }

        [Flags]
        public enum RegOption
        {
            NonVolatile = 0x0,
            Volatile = 0x1,
            CreateLink = 0x2,
            BackupRestore = 0x4,
            OpenLink = 0x8
        }

        [Flags]
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public enum RegSAM
        {
            QueryValue = 0x0001,
            SetValue = 0x0002,
            CreateSubKey = 0x0004,
            EnumerateSubKeys = 0x0008,
            Notify = 0x0010,
            CreateLink = 0x0020,
            WOW64_32Key = 0x0200,
            WOW64_64Key = 0x0100,
            WOW64_Res = 0x0300,
            Read = 0x00020019,
            Write = 0x00020006,
            Execute = 0x00020019,
            AllAccess = 0x000f003f
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        private static extern int RegCreateKeyEx(SafeHandleZeroOrMinusOneIsInvalid hKey, string lpSubKey, int reserved, string lpClass, int dwOptions, int samDesigner, IntPtr lpSecurityAttributes, out IntPtr hkResult, out int lpdwDisposition);
    }
}