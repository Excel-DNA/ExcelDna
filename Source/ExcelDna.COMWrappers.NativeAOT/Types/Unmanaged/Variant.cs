using System.Runtime.InteropServices;

namespace Addin.Types.Unmanaged;

[StructLayout(LayoutKind.Explicit)]
public struct Variant
{
    [FieldOffset(0)]
    public DECIMAL decVal;

    [FieldOffset(0)]
    public ushort vt;

    [FieldOffset(2)]
    public ushort wReserved1;

    [FieldOffset(4)]
    public ushort wReserved2;

    [FieldOffset(6)]
    public ushort wReserved3;

    [FieldOffset(8)]
    public long llVal;

    [FieldOffset(8)]
    public int lVal;

    [FieldOffset(8)]
    public byte bVal;

    [FieldOffset(8)]
    public short iVal;

    [FieldOffset(8)]
    public float fltVal;

    [FieldOffset(8)]
    public double dblVal;

    [FieldOffset(8)]
    public short boolVal;

    [FieldOffset(8)]
    public short __OBSOLETE__VARIANT_BOOL;

    [FieldOffset(8)]
    public int scode;

    [FieldOffset(8)]
    public CY cyVal;

    [FieldOffset(8)]
    public double date;

    [FieldOffset(8)]
    public nint bstrVal;

    [FieldOffset(8)]
    public nint punkVal;

    [FieldOffset(8)]
    public nint pdispVal;

    [FieldOffset(8)]
    public nint parray;

    [FieldOffset(8)]
    public nint pbVal;

    [FieldOffset(8)]
    public nint piVal;

    [FieldOffset(8)]
    public nint plVal;

    [FieldOffset(8)]
    public nint pllVal;

    [FieldOffset(8)]
    public nint pfltVal;

    [FieldOffset(8)]
    public nint pdblVal;

    [FieldOffset(8)]
    public nint pboolVal;

    [FieldOffset(8)]
    public nint __OBSOLETE__VARIANT_PBOOL;

    [FieldOffset(8)]
    public nint pscode;

    [FieldOffset(8)]
    public nint pcyVal;

    [FieldOffset(8)]
    public nint pdate;

    [FieldOffset(8)]
    public nint pbstrVal;

    [FieldOffset(8)]
    public nint ppunkVal;

    [FieldOffset(8)]
    public nint ppdispVal;

    [FieldOffset(8)]
    public nint pparray;

    [FieldOffset(8)]
    public nint pvarVal;

    [FieldOffset(8)]
    public nint byref;

    [FieldOffset(8)]
    public sbyte cVal;

    [FieldOffset(8)]
    public ushort uiVal;

    [FieldOffset(8)]
    public uint ulVal;

    [FieldOffset(8)]
    public ulong ullVal;

    [FieldOffset(8)]
    public int intVal;

    [FieldOffset(8)]
    public uint uintVal;

    [FieldOffset(8)]
    public nint pdecVal;

    [FieldOffset(8)]
    public nint pcVal;

    [FieldOffset(8)]
    public nint puiVal;

    [FieldOffset(8)]
    public nint pulVal;

    [FieldOffset(8)]
    public nint pullVal;

    [FieldOffset(8)]
    public nint pintVal;

    [FieldOffset(8)]
    public nint puintVal;

    [FieldOffset(8)]
    public __tagBRECORD __tagBRECORD;
}

[StructLayout(LayoutKind.Explicit)]
public struct CY
{
    [FieldOffset(0)]
    public uint Lo;

    [FieldOffset(4)]
    public int Hi;

    [FieldOffset(0)]
    public long int64;
}

[StructLayout(LayoutKind.Sequential)]
public struct __tagBRECORD
{
    public nint pvRecord;
    public nint pRecInfo;
}

[StructLayout(LayoutKind.Explicit)]
public struct DECIMAL
{
    [FieldOffset(0)]
    public ushort wReserved;

    [FieldOffset(2)]
    public ushort signscale;

    [FieldOffset(2)]
    public byte scale;

    [FieldOffset(3)]
    public byte sign;

    [FieldOffset(4)]
    public uint Hi32;

    [FieldOffset(8)]
    public ulong Lo64;

    [FieldOffset(8)]
    public uint Lo32;

    [FieldOffset(12)]
    public uint Mid32;
}
