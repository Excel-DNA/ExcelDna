using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Addin.Types.Marshalling;

[CustomMarshaller(typeof(Managed.ExcepInfo), MarshalMode.Default, typeof(ExcepInfo))]
public static class ExcepInfo
{
    public static Unmanaged.ExcepInfo ConvertToUnmanaged(Managed.ExcepInfo managed)
    {
        return new Unmanaged.ExcepInfo
        {
            bstrDescription = Marshal.StringToBSTR(managed.bstrDescription),
            bstrHelpFile = Marshal.StringToBSTR(managed.bstrHelpFile),
            bstrSource = Marshal.StringToBSTR(managed.bstrSource),
            dwHelpContext = managed.dwHelpContext,
            pfnDeferredFillIn = managed.pfnDeferredFillIn,
            pvReserved = managed.pvReserved,
            scode = managed.scode,
            wCode = managed.wCode,
            wReserved = managed.wReserved,
        };
    }

    public static Managed.ExcepInfo ConvertToManaged(Unmanaged.ExcepInfo unmanaged)
    {
        return new Managed.ExcepInfo
        {
            bstrDescription =
                unmanaged.bstrDescription != 0
                    ? Marshal.PtrToStringBSTR(unmanaged.bstrDescription)
                    : "",
            bstrHelpFile =
                unmanaged.bstrHelpFile != 0 ? Marshal.PtrToStringBSTR(unmanaged.bstrHelpFile) : "",
            bstrSource =
                unmanaged.bstrSource != 0 ? Marshal.PtrToStringBSTR(unmanaged.bstrSource) : "",
            dwHelpContext = unmanaged.dwHelpContext,
            pfnDeferredFillIn = unmanaged.pfnDeferredFillIn,
            pvReserved = unmanaged.pvReserved,
            scode = unmanaged.scode,
            wCode = unmanaged.wCode,
            wReserved = unmanaged.wReserved,
        };
    }
}
