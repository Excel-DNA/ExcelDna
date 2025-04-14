using System.Runtime.InteropServices;

namespace ExcelDna.COMWrappers.NativeAOT.ComInterfaces
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct ExcepInfoNative
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
}
