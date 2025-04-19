#if COM_GENERATED

using System.Linq;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [CustomMarshaller(typeof(DispParams), MarshalMode.Default, typeof(DispParamsMarshaller))]
    internal static class DispParamsMarshaller
    {
        public static unsafe DispParamsNative ConvertToUnmanaged(DispParams managed)
        {
            return new DispParamsNative
            {
                cArgs = managed.cArgs,
                cNamedArgs = managed.cNamedArgs,
                rgdispidNamedArgs = managed.rgdispidNamedArgs != 0 ? &managed.rgdispidNamedArgs : null,
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
            };
        }
    }
}

#endif
