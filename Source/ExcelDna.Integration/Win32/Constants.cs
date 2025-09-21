#if !USE_WINDOWS_FORMS

namespace ExcelDna.Integration.Win32
{
    internal class Constants
    {
        public const int ES_MULTILINE = 4;
        public const int ES_AUTOVSCROLL = 0x40;

        public const int WM_SIZE = 5;

        public const int CW_USEDEFAULT = unchecked((int)0x80000000);
    }
}

#endif
