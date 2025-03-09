using System.Runtime.InteropServices;

namespace Addin.Types.Unmanaged;

[StructLayout(LayoutKind.Sequential)]
public struct ExcepInfo
{
    public short wCode;
    public short wReserved;
    public nint bstrSource;
    public nint bstrDescription;
    public nint bstrHelpFile;
    public int dwHelpContext;
    public nint pvReserved;
    public nint pfnDeferredFillIn;
    public int scode;
}
