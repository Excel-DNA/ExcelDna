#if COM_GENERATED

using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [CustomMarshaller(typeof(DispParams), MarshalMode.Default, typeof(DispParamsMarshaller))]
    internal static class DispParamsMarshaller
    {
        public static unsafe DispParamsNative ConvertToUnmanaged(DispParams managed)
        {
            int* rgdispidNamedArgs = null;
            if (managed.rgdispidNamedArgs != 0)
            {
                rgdispidNamedArgs = (int*)Marshal.AllocHGlobal(sizeof(int));
                *rgdispidNamedArgs = managed.rgdispidNamedArgs;
            }

            return new DispParamsNative
            {
                cArgs = managed.cArgs,
                cNamedArgs = managed.cNamedArgs,
                rgdispidNamedArgs = rgdispidNamedArgs,
                rgvarg =
                    managed.rgvarg != null
                        ? ArrayMarshaller.ArrayToPtr(managed.rgvarg.Reverse().Select(VariantMarshaller.ConvertToUnmanaged).ToArray())
                        : nint.Zero
            };
        }

        public static unsafe DispParams ConvertToManaged(DispParamsNative unmanaged)
        {
            return new DispParams
            {
                cArgs = unmanaged.cArgs,
                cNamedArgs = unmanaged.cNamedArgs,
                rgdispidNamedArgs = unmanaged.rgdispidNamedArgs != null ? *unmanaged.rgdispidNamedArgs : 0,
                rgvarg = ArrayMarshaller
                    .PtrToArray<VariantNative>(unmanaged.rgvarg, unmanaged.cArgs)
                    .Select(VariantMarshaller.ConvertToManaged)
                    .Reverse()
                    .ToArray(),
                rgvargNative = unmanaged.rgvarg,
            };
        }

        public static unsafe void Free(DispParamsNative unmanaged)
        {
            if (unmanaged.rgvarg != 0)
            {
                int size = Marshal.SizeOf<VariantNative>();
                for (int i = 0; i < unmanaged.cArgs; ++i)
                {
                    VariantMarshaller.Free(Marshal.PtrToStructure<VariantNative>(unmanaged.rgvarg + i * size));
                }
                ArrayMarshaller.FreePtr(unmanaged.rgvarg);
            }

            if (unmanaged.rgdispidNamedArgs != null)
                Marshal.FreeHGlobal((nint)unmanaged.rgdispidNamedArgs);
        }

        public static void UpdateArg(DispParams dp, Variant v, int i)
        {
            int ri = dp.cArgs - 1 - i;
            var size = Marshal.SizeOf<VariantNative>();
            VariantNative vn = VariantMarshaller.ConvertToUnmanaged(v);
            Marshal.StructureToPtr<VariantNative>(vn, dp.rgvargNative + ri * size, false);
        }

        public static void UpdateRefIntArg(DispParams dp, int v, int i)
        {
            VariantMarshaller.UpdateRefInt(GetVariantParam(dp, i), v);
        }

        public static void UpdateRefBoolArg(DispParams dp, bool v, int i)
        {
            VariantMarshaller.UpdateRefBool(GetVariantParam(dp, i), v);
        }

        private static VariantNative GetVariantParam(DispParams dp, int i)
        {
            int ri = dp.cArgs - 1 - i;
            var size = Marshal.SizeOf<VariantNative>();
            return Marshal.PtrToStructure<VariantNative>(dp.rgvargNative + ri * size);
        }
    }
}

#endif
