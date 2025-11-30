#if COM_GENERATED

using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal static class ArrayMarshaller
    {
        public unsafe static nint ArrayToPtr<T>(T[] str)
        {
            return (nint)Unsafe.AsPointer(ref MemoryMarshal.GetArrayDataReference(str));
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
