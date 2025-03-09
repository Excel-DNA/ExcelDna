using System.Runtime.InteropServices;

namespace Addin.Types.Unmanaged;

[StructLayout(LayoutKind.Sequential)]
public unsafe struct DispParams
{
    public nint rgvarg;
    public int* rgdispidNamedArgs;
    public int cArgs;
    public int cNamedArgs;
}
