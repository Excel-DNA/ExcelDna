#if COM_GENERATED

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    internal partial interface IRibbonControl : IDispatch
    {
        [PreserveSig]
        int get_Id([MarshalAs(UnmanagedType.BStr)] out string result);

        [PreserveSig]
        int get_Context(nint result);

        [PreserveSig]
        int get_Tag([MarshalAs(UnmanagedType.BStr)] out string result);
    }
}

#endif
