namespace ExcelDna.COMWrappers.NativeAOT.ComInterfaces
{
    internal struct Variant
    {
        public Variant(object? value)
        {
            Value = value;
        }

        public object? Value { get; set; }
    }
}
