#if !USE_WINDOWS_FORMS

namespace ExcelDna.Integration.Win32
{
    internal class Util
    {
        public static int LoWord(int i) => i & 0xFFFF;
        public static int HiWord(int i) => i >> 16;
    }
}

#endif
