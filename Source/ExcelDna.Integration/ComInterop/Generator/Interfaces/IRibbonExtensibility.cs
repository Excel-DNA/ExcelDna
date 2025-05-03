#if COM_GENERATED

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid(ExcelDna.ComInterop.ComAPI.gstrIRibbonExtensibility)]
    internal partial interface IRibbonExtensibility : IDispatch
    {
        [PreserveSig]
        int GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID, [MarshalAs(UnmanagedType.BStr)] out string result);
    }
}

#endif
