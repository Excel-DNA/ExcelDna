using System.Runtime.InteropServices;

namespace ExcelDna.COMWrappers.NativeAOT.ComInterfaces
{
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct DispParamsNative
    {
        public nint rgvarg;
        public int* rgdispidNamedArgs;
        public int cArgs;
        public int cNamedArgs;
    }
}
