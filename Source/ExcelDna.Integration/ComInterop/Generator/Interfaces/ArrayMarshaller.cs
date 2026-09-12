#if COM_GENERATED

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal static class ArrayMarshaller
    {
        public static nint ArrayToPtr<T>(T[] values)
        {
            if (values == null || values.Length == 0)
                return nint.Zero;

            int size = Marshal.SizeOf<T>();
            nint result = Marshal.AllocHGlobal(size * values.Length);
            for (int i = 0; i < values.Length; ++i)
                Marshal.StructureToPtr(values[i], result + i * size, false);

            return result;
        }

        public static void FreePtr(nint ptr)
        {
            if (ptr != nint.Zero)
                Marshal.FreeHGlobal(ptr);
        }

        public static T[] PtrToArray<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors | DynamicallyAccessedMemberTypes.NonPublicConstructors)] T>(nint str, int len)
        {
            var size = Marshal.SizeOf<T>();
            var ret = new T[len];
            for (int i = 0; i < len; ++i)
            {
                ret[i] = Marshal.PtrToStructure<T>(nint.Add(str, i * size))!;
            }
            return ret;
        }
    }
}

#endif
