using Addin.ComApi;
using Addin.Types.Unmanaged;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Addin.Types.Marshalling;

[CustomMarshaller(typeof(Managed.Variant), MarshalMode.Default, typeof(Variant))]
public static class Variant
{
    public const int DISP_E_PARAMNOTFOUND = -2147352572;

    public static Unmanaged.Variant ConvertToUnmanaged(Managed.Variant managed)
    {
        if (managed.Value == Type.Missing)
        {
            return new Unmanaged.Variant
            {
                vt = (ushort)VariantType.VT_ERROR,
                scode = DISP_E_PARAMNOTFOUND,
            };
        }

        return managed.Value switch
        {
            bool boolVal
                => new Unmanaged.Variant
                {
                    vt = (ushort)VariantType.VT_BOOL,
                    boolVal = (short)(
                        boolVal ? VariantBool.VARIANT_TRUE : VariantBool.VARIANT_FALSE
                    ),
                },
            int lVal => new Unmanaged.Variant { vt = (ushort)VariantType.VT_I4, lVal = lVal, },
            string bstrVal
                => new Unmanaged.Variant
                {
                    vt = (ushort)VariantType.VT_BSTR,
                    bstrVal = Marshal.StringToBSTR(bstrVal),
                },
            null => new Unmanaged.Variant { vt = (ushort)VariantType.VT_NULL, },
            _ =>
                throw new NotImplementedException(managed.Value.GetType().ToString())
    ,
        };
    }

    public static unsafe Managed.Variant ConvertToManaged(Unmanaged.Variant unmanaged)
    {
        if ((VariantType)unmanaged.vt == VariantType.VT_ERROR && unmanaged.scode == DISP_E_PARAMNOTFOUND)
            return new Managed.Variant(Type.Missing);

        return (VariantType)unmanaged.vt switch
        {
            VariantType.VT_BOOL
                => new Managed.Variant
                {
                    Value = unmanaged.boolVal == (short)VariantBool.VARIANT_TRUE,
                },
            VariantType.VT_I4 => new Managed.Variant { Value = unmanaged.lVal, },
            VariantType.VT_BSTR
                => new Managed.Variant { Value = Marshal.PtrToStringBSTR(unmanaged.bstrVal), },
            VariantType.VT_DISPATCH
                => new Managed.Variant
                {
                    Value = ComInterfaceMarshaller<IDispatch>.ConvertToManaged(
                        (void*)unmanaged.pdispVal
                    ),
                },
            VariantType.VT_EMPTY => new Managed.Variant { },
            VariantType.VT_NULL => new Managed.Variant { },
            _ => throw new NotImplementedException(unmanaged.vt.ToString()),
        };
    }
}
