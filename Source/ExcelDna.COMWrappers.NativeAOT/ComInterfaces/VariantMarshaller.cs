using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.COMWrappers.NativeAOT.ComInterfaces
{
    [CustomMarshaller(typeof(Variant), MarshalMode.Default, typeof(VariantMarshaller))]
    internal static class VariantMarshaller
    {
        public const int DISP_E_PARAMNOTFOUND = -2147352572;

        public static VariantNative ConvertToUnmanaged(Variant managed)
        {
            if (managed.Value == Type.Missing)
            {
                return new VariantNative
                {
                    vt = (ushort)VariantTypeNative.VT_ERROR,
                    scode = DISP_E_PARAMNOTFOUND,
                };
            }

            if (managed.Value != null && managed.Value.GetType().IsEnum)
            {
                return new VariantNative { vt = (ushort)VariantTypeNative.VT_I4, lVal = (int)managed.Value };
            }

            return managed.Value switch
            {
                bool boolVal
                    => new VariantNative
                    {
                        vt = (ushort)VariantTypeNative.VT_BOOL,
                        boolVal = (short)(
                            boolVal ? VariantBoolNative.VARIANT_TRUE : VariantBoolNative.VARIANT_FALSE
                        ),
                    },
                int lVal => new VariantNative { vt = (ushort)VariantTypeNative.VT_I4, lVal = lVal, },
                string bstrVal
                    => new VariantNative
                    {
                        vt = (ushort)VariantTypeNative.VT_BSTR,
                        bstrVal = Marshal.StringToBSTR(bstrVal),
                    },
                null => new VariantNative { vt = (ushort)VariantTypeNative.VT_NULL, },
                _ =>
                    throw new NotImplementedException(managed.Value.GetType().ToString())
        ,
            };
        }

        public static unsafe Variant ConvertToManaged(VariantNative unmanaged)
        {
            if ((VariantTypeNative)unmanaged.vt == VariantTypeNative.VT_ERROR && unmanaged.scode == DISP_E_PARAMNOTFOUND)
                return new Variant(Type.Missing);

            return (VariantTypeNative)unmanaged.vt switch
            {
                VariantTypeNative.VT_BOOL
                    => new Variant
                    {
                        Value = unmanaged.boolVal == (short)VariantBoolNative.VARIANT_TRUE,
                    },
                VariantTypeNative.VT_I4 => new Variant { Value = unmanaged.lVal, },
                VariantTypeNative.VT_BSTR
                    => new Variant { Value = Marshal.PtrToStringBSTR(unmanaged.bstrVal), },
                VariantTypeNative.VT_DISPATCH
                    => new Variant
                    {
                        Value = new DispatchObject(unmanaged.pdispVal)
                    },
                VariantTypeNative.VT_EMPTY => new Variant { },
                VariantTypeNative.VT_NULL => new Variant { },
                _ => throw new NotImplementedException(unmanaged.vt.ToString()),
            };
        }
    }
}
