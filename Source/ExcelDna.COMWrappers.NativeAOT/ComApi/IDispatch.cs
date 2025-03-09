using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;
using Managed = Addin.Types.Managed;
using Marshalling = Addin.Types.Marshalling;

namespace Addin.ComApi;

[GeneratedComInterface]
[Guid("00020400-0000-0000-C000-000000000046")] // The IID for IDispatch
public partial interface IDispatch
{
    [PreserveSig]
    int GetTypeInfoCount(out uint pctinfo);

    [PreserveSig]
    int GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo);

    [PreserveSig]
    int GetIDsOfNames(
        ref Guid riid,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
        uint cNames,
        uint lcid,
        [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
    );

    [PreserveSig]
    int Invoke(
        int dispIdMember,
        Guid riid,
        uint lcid,
        INVOKEKIND wFlags,
        [MarshalUsing(typeof(Marshalling.DispParams))] ref Managed.DispParams pDispParams,
        [MarshalUsing(typeof(Marshalling.Variant))] ref Managed.Variant pVarResult,
        [MarshalUsing(typeof(Marshalling.ExcepInfo))] ref Managed.ExcepInfo pExcepInfo,
        ref uint puArgErr
    );
}
