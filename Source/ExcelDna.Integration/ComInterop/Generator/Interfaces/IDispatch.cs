#if COM_GENERATED

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [GeneratedComInterface]
    [Guid("00020400-0000-0000-C000-000000000046")]
    internal partial interface IDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount(out uint pctinfo);

        [PreserveSig]
        int GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo);

        [PreserveSig]
        int GetIDsOfNames(
            Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] rgszNames,
            uint cNames,
            uint lcid,
            [In][Out][MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] int[] rgDispId
        );

        [PreserveSig]
        int Invoke(
            int dispIdMember,
            Guid riid,
            uint lcid,
            INVOKEKIND wFlags,
            [MarshalUsing(typeof(DispParamsMarshaller))] ref DispParams pDispParams,
            [MarshalUsing(typeof(VariantMarshaller))] out Variant pVarResult,
            [MarshalUsing(typeof(ExcepInfoMarshaller))] out ExcepInfo pExcepInfo,
            ref uint puArgErr
        );
    }
}

#endif
