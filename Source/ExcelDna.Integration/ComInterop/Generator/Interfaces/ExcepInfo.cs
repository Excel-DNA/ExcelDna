#if COM_GENERATED

namespace ExcelDna.Integration.ComInterop.Generator.Interfaces
{
    internal struct ExcepInfo
    {
        //
        // Summary:
        //     Describes the error intended for the customer.
        public string bstrDescription;

        //
        // Summary:
        //     Contains the fully-qualified drive, path, and file name of a Help file that contains
        //     more information about the error.
        public string bstrHelpFile;

        //
        // Summary:
        //     Indicates the name of the source of the exception. Typically, this is an application
        //     name.
        public string bstrSource;

        //
        // Summary:
        //     Indicates the Help context ID of the topic within the Help file.
        public int dwHelpContext;

        //
        // Summary:
        //     Represents a pointer to a function that takes an System.Runtime.InteropServices.EXCEPINFO
        //     structure as an argument and returns an HRESULT value. If deferred fill-in is
        //     not desired, this field is set to null.
        public nint pfnDeferredFillIn;

        //
        // Summary:
        //     This field is reserved; it must be set to null.
        public nint pvReserved;

        //
        // Summary:
        //     A return value describing the error.
        public int scode;

        //
        // Summary:
        //     Represents an error code identifying the error.
        public short wCode;

        //
        // Summary:
        //     This field is reserved; it must be set to 0.
        public short wReserved;
    }
}

#endif
