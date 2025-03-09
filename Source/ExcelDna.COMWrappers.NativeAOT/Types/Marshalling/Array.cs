using System.Runtime.InteropServices;

namespace Addin.Types.Marshalling;

public static class Array
{
    public static nint ArrayToPtr<T>(T[] str)
    {
        var size = Marshal.SizeOf<T>();
        var len = str.Length;
        var ptrs = new nint[str.Length];
        var basePtr = Marshal.AllocHGlobal(size * len);
        for (int i = 0; i < len; ++i)
        {
            ptrs[i] = nint.Add(basePtr, i * size);
            Marshal.StructureToPtr(str[i], ptrs[i], false);
        }
        return basePtr;
    }

    public static T[] PtrToArray<T>(nint str, int len)
    {
        var size = Marshal.SizeOf<T>();
        var ret = new T[len];
        for (int i = 0; i < len; ++i)
        {
            ret[i] = Marshal.PtrToStructure<T>(nint.Add(str, i * size));
        }
        return ret;
    }
}
