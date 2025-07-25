﻿using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using CLSID = System.Guid;

namespace ExcelDna.Integration.ComInterop
{
    internal class TypeAdapter : IType
    {
        private static readonly CultureInfo _enUsCulture = new CultureInfo(1033);

        public object GetProperty(string name, object comObject)
        {
            return comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, null, _enUsCulture);
        }

        public object GetIndex(int i, object comObject)
        {
            return comObject.GetType().InvokeMember("", BindingFlags.GetProperty, null, comObject, new object[] { i }, _enUsCulture);
        }

        public object GetIndex(string name, object comObject)
        {
            return comObject.GetType().InvokeMember("", BindingFlags.GetProperty, null, comObject, new object[] { name }, _enUsCulture);
        }

        public bool Is(CLSID guid, object comObject)
        {
            IntPtr pUnk = Marshal.GetIUnknownForObject(comObject);

            IntPtr pObj;
            Marshal.QueryInterface(pUnk, ref guid, out pObj);
            return (pObj != IntPtr.Zero);
        }

        public object Invoke(string name, object[] args, object comObject)
        {
            return comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, comObject, args, _enUsCulture);
        }

        public void SetProperty(string name, object value, object comObject)
        {
            comObject.GetType().InvokeMember(name, BindingFlags.SetProperty, null, comObject, new object[] { value }, _enUsCulture);
        }

        public object GetObject(IntPtr pUnk)
        {
            return Marshal.GetObjectForIUnknown(pUnk);
        }

        public void ReleaseObject(object comObject)
        {
            Marshal.ReleaseComObject(comObject);
        }

        public bool HasProperty(string name, object comObject)
        {
            return ExcelDna.ComInterop.DispatchHelper.HasProperty(comObject, name);
        }

        public IntPtr QueryInterface(CLSID guid, object comObject)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(comObject);
            if (Marshal.QueryInterface(punk, ref guid, out IntPtr result) == 0)
            {
                return result;
            }

            return IntPtr.Zero;
        }
    }
}
