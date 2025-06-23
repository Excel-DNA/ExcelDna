#if COM_GENERATED

using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
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

#endif
