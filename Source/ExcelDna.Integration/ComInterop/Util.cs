namespace ExcelDna.Integration.ComInterop
{
    internal static class Util
    {
        private static TypeAdapter typeAdapter = new TypeAdapter();

        public static IType TypeAdapter
        {
            get
            {
                return NativeAOT.IsActive ? NativeAOT.TypeAdapter : typeAdapter;
            }
        }
    }
}
