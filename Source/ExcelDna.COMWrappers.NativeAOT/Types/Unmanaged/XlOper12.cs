using System.Runtime.InteropServices;

namespace Addin.Types.Unmanaged;

[StructLayout(LayoutKind.Sequential)]
public struct xlref12
{
    /// RW->INT32->int
    public int rwFirst;

    /// RW->INT32->int
    public int rwLast;

    /// COL->INT32->int
    public int colFirst;

    /// COL->INT32->int
    public int colLast;
}

[StructLayout(LayoutKind.Sequential)]
public struct xmlref12
{
    /// WORD->short
    public short count;

    /// XLREF12[1]
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1, ArraySubType = UnmanagedType.Struct)]
    public xlref12[] reftbl;
}

[StructLayout(LayoutKind.Sequential)]
public struct FP12
{
    /// INT32->int
    public int rows;

    /// INT32->int
    public int columns;

    /// double[1]
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1, ArraySubType = UnmanagedType.R8)]
    public double[] array;
}

[StructLayout(LayoutKind.Sequential)]
public struct sref
{
    /// WORD->short
    public short count;

    /// XLREF12->xlref12
    public xlref12 @ref;
}

[StructLayout(LayoutKind.Sequential)]
public struct mref
{
    /// XLMREF12*
    public nint lpmref;

    /// IDSHEET->DWORD_PTR->ULONG_PTR->int
    public int idSheet;
}

[StructLayout(LayoutKind.Sequential)]
public struct array
{
    /// xloper12*
    public nint lparray;

    /// RW->INT32->int
    public int rows;

    /// COL->INT32->int
    public int columns;
}

[StructLayout(LayoutKind.Explicit)]
public struct valflow
{
    /// int
    [FieldOffset(0)]
    public int level;

    /// int
    [FieldOffset(0)]
    public int tbctrl;

    /// IDSHEET->DWORD_PTR->ULONG_PTR->int
    [FieldOffset(0)]
    public int idSheet;
}

[StructLayout(LayoutKind.Sequential)]
public struct flow
{
    public valflow valflow;

    /// RW->INT32->int
    public int rw;

    /// COL->INT32->int
    public int col;

    /// BYTE->char
    public byte xlflow;
}

[StructLayout(LayoutKind.Explicit)]
public struct h
{
    /// BYTE*
    [FieldOffset(0)]
    public nint lpbData;

    /// HANDLE->void*
    [FieldOffset(0)]
    public nint hdata;
}

[StructLayout(LayoutKind.Sequential)]
public struct bigdata
{
    public h h;

    /// int
    public int cbData;
}

[StructLayout(LayoutKind.Explicit)]
public struct val
{
    /// double
    [FieldOffset(0)]
    public double num;

    /// XCHAR*
    [FieldOffset(0)]
    public nint str;

    /// BOOL->INT32->int
    [FieldOffset(0)]
    [MarshalAs(UnmanagedType.I1)]
    public bool xbool;

    /// int
    [FieldOffset(0)]
    public int err;

    /// int
    [FieldOffset(0)]
    public int w;

    [FieldOffset(0)]
    public sref sref;

    [FieldOffset(0)]
    public mref mref;

    [FieldOffset(0)]
    public array array;

    [FieldOffset(0)]
    public flow flow;

    [FieldOffset(0)]
    public bigdata bigdata;
}

[StructLayout(LayoutKind.Sequential)]
public struct xloper12
{
    public val val;

    /// DWORD->int
    public int xltype;
}
