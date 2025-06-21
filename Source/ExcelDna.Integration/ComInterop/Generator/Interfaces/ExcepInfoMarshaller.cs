#if COM_GENERATED

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    [CustomMarshaller(typeof(ExcepInfo), MarshalMode.Default, typeof(ExcepInfoMarshaller))]
    internal static class ExcepInfoMarshaller
    {
        public static ExcepInfoNative ConvertToUnmanaged(ExcepInfo managed)
        {
            return new ExcepInfoNative
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

        public static ExcepInfo ConvertToManaged(ExcepInfoNative unmanaged)
        {
            return new ExcepInfo
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
}

#endif
