#if COM_GENERATED

using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [CustomMarshaller(typeof(Variant), MarshalMode.Default, typeof(VariantMarshaller))]
    internal static class VariantMarshaller
    {
        private const VariantTypeNative VT_BYREF_BOOL = (VariantTypeNative)(ushort)VariantTypeNative.VT_BOOL + (ushort)VariantTypeNative.VT_BYREF;
        private const VariantTypeNative VT_BYREF_I4 = (VariantTypeNative)(ushort)VariantTypeNative.VT_I4 + (ushort)VariantTypeNative.VT_BYREF;
        private const VariantTypeNative VT_VARIANT_ARRAY = (VariantTypeNative)(ushort)VariantTypeNative.VT_VARIANT + (ushort)VariantTypeNative.VT_ARRAY;

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
                double dVal => new VariantNative { vt = (ushort)VariantTypeNative.VT_R8, dblVal = dVal, },
                string bstrVal
                    => new VariantNative
                    {
                        vt = (ushort)VariantTypeNative.VT_BSTR,
                        bstrVal = Marshal.StringToBSTR(bstrVal),
                    },
                Array array => VariantArrayToUnmanaged(array),
                DispatchObject doVal => new VariantNative { vt = (ushort)VariantTypeNative.VT_DISPATCH, pdispVal = doVal.P, },
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
                VT_BYREF_BOOL => RefBoolToManaged(unmanaged.pboolVal),
                VariantTypeNative.VT_I4 => new Variant { Value = unmanaged.lVal, },
                VT_BYREF_I4 => RefIntToManaged(unmanaged.plVal),
                VariantTypeNative.VT_BSTR
                    => new Variant { Value = Marshal.PtrToStringBSTR(unmanaged.bstrVal), },
                VariantTypeNative.VT_DISPATCH
                    => new Variant
                    {
                        Value = new DispatchObject(unmanaged.pdispVal)
                    },
                VariantTypeNative.VT_EMPTY => new Variant { },
                VT_VARIANT_ARRAY => VariantArrayToManaged(unmanaged.parray),
                VariantTypeNative.VT_NULL => new Variant { },
                _ => throw new NotImplementedException(unmanaged.vt.ToString()),
            };
        }

        private static Variant VariantArrayToManaged(nint parray)
        {
            SafeArray sa = Marshal.PtrToStructure<SafeArray>(parray);
            if (sa.cDims == 1)
            {
                VariantNative[] vna = ArrayMarshaller.PtrToArray<VariantNative>(sa.pvData, (int)sa.rgsabound.Data.cElements);
                return new Variant(ArrayMarshaller.PtrToArray<VariantNative>(sa.pvData, (int)sa.rgsabound.Data.cElements).Select(i => ConvertToManaged(i)).ToArray());
            }

            throw new NotImplementedException();
        }

        private static VariantNative VariantArrayToUnmanaged(Array array)
        {
            if (array.Rank != 2)
                throw new NotImplementedException();

            nint pBounds = Marshal.AllocHGlobal(Marshal.SizeOf<SAFEARRAYBOUND>() * array.Rank);
            Marshal.StructureToPtr(new SAFEARRAYBOUND { cElements = (uint)array.GetLength(0), lLbound = 0 }, pBounds, false);
            Marshal.StructureToPtr(new SAFEARRAYBOUND { cElements = (uint)array.GetLength(1), lLbound = 0 }, pBounds + Marshal.SizeOf<SAFEARRAYBOUND>(), false);

            nint psa = SafeArray.SafeArrayCreate((ushort)VariantTypeNative.VT_VARIANT, (uint)array.Rank, pBounds);
            SafeArray sa = Marshal.PtrToStructure<SafeArray>(psa);

            for (int col = 0; col < array.GetLength(1); ++col)
            {
                for (int row = 0; row < array.GetLength(0); ++row)
                {
                    int i = col * array.GetLength(0) + row;
                    Marshal.StructureToPtr(ConvertToUnmanaged(new Variant(array.GetValue(row, col))), sa.pvData + i * (int)sa.cbElements, false);
                }
            }

            return new VariantNative { vt = (ushort)VT_VARIANT_ARRAY, parray = psa };
        }

        private static Variant RefBoolToManaged(nint pboolVal)
        {
            short boolVal = Marshal.PtrToStructure<short>(pboolVal);
            return new Variant(boolVal == (short)VariantBoolNative.VARIANT_TRUE);
        }

        private static Variant RefIntToManaged(nint plVal)
        {
            return new Variant(Marshal.PtrToStructure<int>(plVal));
        }

        public static void UpdateRefInt(VariantNative unmanaged, int v)
        {
            Marshal.StructureToPtr(v, unmanaged.plVal, false);
        }
    }
}

#endif
