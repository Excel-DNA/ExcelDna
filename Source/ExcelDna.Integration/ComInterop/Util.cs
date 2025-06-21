﻿using System;
using System.Runtime.InteropServices;
#if COM_GENERATED
using System.Runtime.InteropServices.Marshalling;
#endif

namespace ExcelDna.Integration.ComInterop
{
    internal static class Util
    {
        private static TypeAdapter typeAdapter = new TypeAdapter();
#if COM_GENERATED
        private static ComWrappers comWrappers = new StrategyBasedComWrappers();
        private static Generator.TypeAdapter generatorAdapter = new();
#endif

        public static IType TypeAdapter
        {
            get
            {
#if COM_GENERATED
                return NativeAOT.IsActive ? generatorAdapter : typeAdapter;
#else
                return typeAdapter;
#endif
            }
        }

        public static int QueryInterfaceForObject(object o, Guid guid, out IntPtr ppv)
        {
            Guid iid = guid;
            return Marshal.QueryInterface(GetIUnknownForObject(o), ref iid, out ppv);
        }

        private static IntPtr GetIUnknownForObject(object o)
        {
#if COM_GENERATED
            Type t = o.GetType();
            if (t == typeof(ComObject) || t.GetCustomAttributes(typeof(GeneratedComClassAttribute), false).Length > 0)
            {
                return comWrappers.GetOrCreateComInterfaceForObject(o, CreateComInterfaceFlags.None);
            }
            else
#endif
            {
                return Marshal.GetIUnknownForObject(o);
            }
        }
    }
}
