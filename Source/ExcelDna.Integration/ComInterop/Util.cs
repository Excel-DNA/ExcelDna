namespace ExcelDna.Integration.ComInterop
{
    internal static class Util
    {
        private static TypeAdapter typeAdapter = new TypeAdapter();

#if COM_GENERATED
        private static Generator.TypeAdapter generatorAdapter = new();
#endif

        public static IType TypeAdapter
        {
            get
            {
#if COM_GENERATED
                return NativeAOT.IsActive ? generatorAdapter : typeAdapter;
#else
                return typeAdapter;
#endif
            }
        }
    }
}
