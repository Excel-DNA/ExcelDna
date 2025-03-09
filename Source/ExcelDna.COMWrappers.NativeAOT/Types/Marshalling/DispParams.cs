using System.Runtime.InteropServices.Marshalling;

namespace Addin.Types.Marshalling;

[CustomMarshaller(typeof(Managed.DispParams), MarshalMode.Default, typeof(DispParams))]
public static class DispParams
{
    public static unsafe Unmanaged.DispParams ConvertToUnmanaged(Managed.DispParams managed)
    {
        return new Unmanaged.DispParams
        {
            cArgs = managed.cArgs,
            cNamedArgs = managed.cNamedArgs,
            rgdispidNamedArgs = &managed.rgdispidNamedArgs,
            rgvarg =
                managed.rgvarg != null
                    ? Array.ArrayToPtr(managed.rgvarg.Select(Variant.ConvertToUnmanaged).ToArray())
                    : nint.Zero
        };
    }

    public static unsafe Managed.DispParams ConvertToManaged(Unmanaged.DispParams unmanaged)
    {
        return new Managed.DispParams
        {
            cArgs = unmanaged.cArgs,
            cNamedArgs = unmanaged.cNamedArgs,
            rgdispidNamedArgs = *unmanaged.rgdispidNamedArgs,
            rgvarg = Array
                .PtrToArray<Unmanaged.Variant>(unmanaged.rgvarg, unmanaged.cArgs)
                .Select(Variant.ConvertToManaged)
                .ToArray(),
        };
    }
}
